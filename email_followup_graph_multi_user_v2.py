"""
Email Follow-up Automation using Microsoft Graph API
Analyzes sent emails from the past 7 days and identifies those that require replies
but haven't received them within 2 days.
"""

import os
import sys
import json
import io

# Fix Windows console encoding for emojis
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')


class TeeOutput:
    """Class to duplicate stdout to both console and a file."""
    def __init__(self, log_file_path):
        self.terminal = sys.stdout
        self.log_file = open(log_file_path, 'w', encoding='utf-8')
    
    def write(self, message):
        self.terminal.write(message)
        self.log_file.write(message)
        self.log_file.flush()  # Ensure immediate write
    
    def flush(self):
        self.terminal.flush()
        self.log_file.flush()
    
    def close(self):
        self.log_file.close()
import datetime as dt
from typing import List, Dict, Optional, Tuple
from dateutil import parser as date_parser
from dateutil import tz
import html
import re
from urllib.parse import quote

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("Warning: python-dotenv not installed. Make sure to set environment variables manually.")

try:
    from msal import ConfidentialClientApplication, PublicClientApplication
    import requests
except ImportError:
    print("Please install required packages: pip install msal requests")
    sys.exit(1)

try:
    import openai
except ImportError:
    print("Please install openai package: pip install openai")
    sys.exit(1)
    
# Add this import after the openai import (around line 38)
try:
    from exclusion_checker import is_email_instance_excluded
    EXCLUSION_CHECKER_AVAILABLE = True
except ImportError:
    print("Warning: exclusion_checker module not found. 'Mark as Dealt With' feature disabled.")
    EXCLUSION_CHECKER_AVAILABLE = False

# Configuration
LOOKBACK_DAYS = 7
REPLY_WAIT_DAYS = 2
RECENT_THRESHOLD_DAYS = 2  # Emails newer than this are "recent"
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
LOCAL_TZ = tz.tzlocal()

# Test limit - set to None for no limit, or a number to limit emails processed
TEST_LIMIT = None  # No limit - process all emails

# Priority Keywords for detection
URGENT_KEYWORDS = [
    "urgent", "asap", "immediately", "critical", "emergency", "time-sensitive",
    "high priority", "top priority", "matter of urgency"
]
DEADLINE_KEYWORDS = [
    "deadline", "eod", "end of day", "by tomorrow", "due date", "by friday",
    "by monday", "by end of week", "today", "tonight", "this week", "time bound"
]
ACTION_KEYWORDS = [
    "action required", "please confirm", "need your approval", "waiting for your",
    "please review", "need your input", "requires your attention", "please respond",
    "awaiting your", "pending your"
]

# Environment variables
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")  # For app-only auth (LESS SECURE - see below)
TENANT_ID = os.getenv("AZURE_TENANT_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
USER_EMAIL = os.getenv("USER_EMAIL")  # Only used for app-only auth (ignored in delegated mode)
# Webhook API URL for "Mark as Dealt With" feature
WEBHOOK_API_URL = os.getenv("WEBHOOK_API_URL", "http://localhost:5000")

# For interactive auth (if not using app-only)
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else "https://login.microsoftonline.com/common"
SCOPES = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/Mail.ReadWrite"]

# SECURITY MODEL:
# ===============
# Option 1: DELEGATED AUTH (RECOMMENDED - More Secure)
#   - Do NOT set CLIENT_SECRET in .env
#   - User must sign in via browser each time
#   - Script can ONLY access the signed-in user's mailbox
#   - USER_EMAIL in .env is IGNORED - uses whoever logs in
#
# Option 2: APP-ONLY AUTH (Less Secure - for automation only)
#   - Set CLIENT_SECRET in .env
#   - No user sign-in required
#   - ⚠️ WARNING: Can access ANY mailbox in tenant!
#   - Should only be used in secured environments (servers, CI/CD)
#   - Consider using Exchange Application Access Policies to restrict

USE_DELEGATED_AUTH = not bool(CLIENT_SECRET)

class GraphAPIClient:
    """Microsoft Graph API client for accessing Outlook mailbox."""
    
    def __init__(self, client_id: str, client_secret: Optional[str] = None, tenant_id: Optional[str] = None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}" if tenant_id else "https://login.microsoftonline.com/common"
        self.access_token = None
        self.authenticated_user_email = None  # Set after authentication (delegated auth only)
        self.app = None
        
        if client_secret:
            # App-only authentication (client credentials flow)
            # ⚠️ SECURITY WARNING: This mode can access ANY mailbox in the tenant!
            self.app = ConfidentialClientApplication(
                client_id=client_id,
                client_credential=client_secret,
                authority=self.authority
            )
        else:
            # Interactive browser authentication - MORE SECURE & ENTERPRISE-FRIENDLY
            # User must sign in via browser, and script can only access THEIR mailbox
            self.app = PublicClientApplication(
                client_id=client_id,
                authority=self.authority
            )
    
    def get_access_token(self) -> str:
        """Get access token for Microsoft Graph API."""
        if self.client_secret:
            # Client credentials flow (app-only)
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        else:
            # Interactive browser authentication - MORE COMPATIBLE with enterprise policies
            # First, try to get cached token
            accounts = self.app.get_accounts()
            if accounts:
                result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    # Get user email from cached account
                    self.authenticated_user_email = accounts[0].get("username")
                    return self.access_token
            
            # No cached token, need interactive login via browser
            print("\n🌐 Opening browser for authentication...")
            print("Please sign in with your Microsoft account in the browser window.")
            print("If you see an 'access_denied' error, check the error details in the browser.")
            result = self.app.acquire_token_interactive(
                scopes=SCOPES,
                prompt="consent",  # Force consent screen to ensure permissions are granted
                # Use localhost redirect URI (must match Azure App Registration)
                # If this fails, add "http://localhost" as a Redirect URI in Azure Portal
                port=None  # Let MSAL choose an available port
            )
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            
            # For delegated auth, extract authenticated user's email
            if not self.client_secret:
                # Get user info from the token claims or /me endpoint
                try:
                    me_info = self.make_request("GET", "/me?$select=mail,userPrincipalName")
                    self.authenticated_user_email = me_info.get("mail") or me_info.get("userPrincipalName")
                    print(f"\n✅ Authenticated as: {self.authenticated_user_email}")
                except Exception as e:
                    print(f"Warning: Could not get user info: {e}")
            
            return self.access_token
        else:
            # Enhanced error reporting for troubleshooting
            error_code = result.get("error", "Unknown error")
            error_description = result.get("error_description", "No description provided")
            error_uri = result.get("error_uri", "")
            correlation_id = result.get("correlation_id", "")
            
            error_msg = f"{error_code}"
            if error_description and error_description != error_code:
                error_msg += f": {error_description}"
            if error_uri:
                error_msg += f" ({error_uri})"
            if correlation_id:
                error_msg += f" [Correlation ID: {correlation_id}]"
            
            # Provide helpful guidance based on error type
            if "access_denied" in error_code.lower():
                error_msg += "\n\n💡 Troubleshooting 'access_denied' error:"
                error_msg += "\n   1. Check if admin consent is granted (Azure Portal → App Registration → API permissions)"
                error_msg += "\n   2. Verify 'Enabled for users to sign-in?' is Yes (Enterprise Applications → Properties)"
                error_msg += "\n   3. Check Conditional Access Policies - your IT admin may need to exclude this app"
                error_msg += "\n   4. Ensure the app redirect URI 'http://localhost' is configured in Azure"
            
            raise Exception(f"Failed to acquire token: {error_msg}")
    
    def make_request(self, method: str, endpoint: str, **kwargs) -> dict:
        """Make a request to Microsoft Graph API."""
        if not self.access_token:
            self.get_access_token()
        
        url = f"{GRAPH_API_ENDPOINT}{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        response = requests.request(method, url, headers=headers, **kwargs)
        
        # Provide better error messages
        if not response.ok:
            error_detail = ""
            try:
                error_json = response.json()
                error_detail = error_json.get("error", {}).get("message", "")
            except:
                error_detail = response.text[:200]
            
            error_msg = f"HTTP {response.status_code}: {response.reason}"
            if error_detail:
                error_msg += f" - {error_detail}"
            
            # Special handling for /me endpoint with app-only auth
            if "/me" in endpoint and response.status_code == 400:
                error_msg += "\n\nNote: The /me endpoint doesn't work with app-only authentication. "
                error_msg += "Please set USER_EMAIL in your .env file when using client credentials flow."
            
            raise Exception(error_msg)
        
        return response.json()
    
    def get_user_info(self, user_email: Optional[str] = None) -> dict:
        """
        Get user information.
        For app-only auth, user_email is required.
        For delegated auth, user_email is optional (uses /me if not provided).
        """
        if user_email:
            return self.make_request("GET", f"/users/{user_email}")
        else:
            # Try /me endpoint (only works with delegated permissions)
            try:
                return self.make_request("GET", "/me")
            except Exception as e:
                raise Exception(
                    f"Cannot use /me endpoint. This usually means you're using app-only authentication. "
                    f"Please set USER_EMAIL in your .env file. Original error: {e}"
                )
    
    def get_sent_messages(self, user_email: Optional[str] = None, since_date: dt.datetime = None) -> List[dict]:
        """Get sent messages from the specified user's mailbox."""
        if not user_email:
            # Try /me endpoint (only works with delegated permissions)
            user_endpoint = "/me"
        else:
            user_endpoint = f"/users/{user_email}"
        
        # Build filter for sent items in the past N days
        if since_date:
            since_str = since_date.strftime("%Y-%m-%dT%H:%M:%SZ")
            filter_query = f"sentDateTime ge {since_str}"
        else:
            since_date = dt.datetime.now(tz=tz.UTC) - dt.timedelta(days=LOOKBACK_DAYS)
            since_str = since_date.strftime("%Y-%m-%dT%H:%M:%SZ")
            filter_query = f"sentDateTime ge {since_str}"
        
        endpoint = f"{user_endpoint}/mailFolders('SentItems')/messages"
        params = {
            "$filter": filter_query,
            "$orderby": "sentDateTime desc",
            "$select": "id,subject,body,bodyPreview,sentDateTime,receivedDateTime,toRecipients,ccRecipients,bccRecipients,from,sender,conversationId,webLink,internetMessageId,conversationIndex,parentFolderId",
            "$top": 999
        }
        
        all_messages = []
        while True:
            result = self.make_request("GET", endpoint, params=params)
            messages = result.get("value", [])
            all_messages.extend(messages)
            
            # Check for next page
            next_link = result.get("@odata.nextLink")
            if not next_link:
                break
            
            # Extract endpoint from next_link
            endpoint = next_link.replace(GRAPH_API_ENDPOINT, "")
            params = {}
        
        return all_messages
    
    def get_conversation_messages_by_subject(self, subject: str, conversation_id: str, user_email: Optional[str] = None) -> List[dict]:
        """
        Get all messages in a conversation by searching for matching subject.
        This is MORE RELIABLE than filtering by conversationId because Microsoft's
        Graph API has known issues with conversationId filtering.
        
        Uses $search which is much more reliable than $filter for this use case.
        """
        if not user_email:
            user_endpoint = "/me"
        else:
            user_endpoint = f"/users/{user_email}"
        
        all_messages = []
        
        # Clean subject for search - remove Re:, Fw:, FW:, RE: prefixes
        clean_subject = subject
        for prefix in ["Re: ", "RE: ", "Fw: ", "FW: ", "Fwd: ", "FWD: "]:
            if clean_subject.startswith(prefix):
                clean_subject = clean_subject[len(prefix):]
        
        # Escape quotes in subject for search
        search_subject = clean_subject.replace('"', '\\"')
        
        # Search in both Inbox and Sent Items
        folders_to_check = ["Inbox", "SentItems"]
        
        for folder in folders_to_check:
            try:
                folder_endpoint = f"{user_endpoint}/mailFolders('{folder}')/messages"
                
                # Use $search which is more reliable than $filter for subject matching
                params = {
                    "$search": f'"subject:{search_subject}"',
                    "$select": "id,subject,body,bodyPreview,sentDateTime,receivedDateTime,toRecipients,ccRecipients,bccRecipients,from,sender,conversationId,isRead,parentFolderId,webLink",
                    "$top": 50
                }
                
                result = self.make_request("GET", folder_endpoint, params=params)
                folder_messages = result.get("value", [])
                
                # Filter to only include messages with matching conversationId
                # (search might return partial matches, so we verify)
                for msg in folder_messages:
                    if msg.get("conversationId") == conversation_id:
                        all_messages.append(msg)
                
            except Exception as e:
                # If search fails for this folder, continue to next
                pass
        
        # Remove duplicates by message ID
        seen_ids = set()
        unique_messages = []
        for msg in all_messages:
            msg_id = msg.get("id")
            if msg_id and msg_id not in seen_ids:
                seen_ids.add(msg_id)
                unique_messages.append(msg)
        
        # Sort by date
        unique_messages.sort(key=lambda x: x.get("sentDateTime") or x.get("receivedDateTime") or "")
        
        return unique_messages
    
    def get_conversation_messages(self, conversation_id: str, user_email: Optional[str] = None, subject_hint: str = None) -> List[dict]:
        """
        Get all messages in a conversation across all folders.
        
        Strategy:
        1. If subject_hint provided, use subject-based search (most reliable)
        2. Otherwise, fetch from Inbox and SentItems and filter by conversationId in code
           (avoids the problematic OData filter on conversationId)
        """
        if not user_email:
            user_endpoint = "/me"
        else:
            user_endpoint = f"/users/{user_email}"
        
        # If we have a subject hint, use the more reliable subject-based search
        if subject_hint:
            messages = self.get_conversation_messages_by_subject(subject_hint, conversation_id, user_email)
            if messages:
                return messages
        
        # Fallback: Fetch recent messages from multiple folders, filter by conversationId in code
        # This avoids the problematic OData filter entirely
        # IMPROVED: Check more folders, longer date range, more messages
        all_messages = []
        folders_to_check = ["Inbox", "SentItems", "Archive", "DeletedItems"]
        
        # Increased from 30 to 90 days to catch more recent activity
        since_date = dt.datetime.now(tz=tz.UTC) - dt.timedelta(days=90)
        since_str = since_date.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        for folder in folders_to_check:
            try:
                folder_endpoint = f"{user_endpoint}/mailFolders('{folder}')/messages"
                
                # Use both sentDateTime and receivedDateTime to catch all messages
                params = {
                    "$filter": f"sentDateTime ge {since_str} or receivedDateTime ge {since_str}",
                    "$select": "id,subject,body,bodyPreview,sentDateTime,receivedDateTime,toRecipients,ccRecipients,bccRecipients,from,sender,conversationId,isRead,parentFolderId,webLink",
                    "$orderby": "sentDateTime desc,receivedDateTime desc",
                    "$top": 500  # Increased from 200 to 500
                }
                
                result = self.make_request("GET", folder_endpoint, params=params)
                folder_messages = result.get("value", [])
                
                # Filter by conversationId IN CODE (not in API filter - that's what fails!)
                for msg in folder_messages:
                    if msg.get("conversationId") == conversation_id:
                        all_messages.append(msg)
                
            except Exception as e:
                # If this folder fails, continue to next
                pass
        
        # Remove duplicates by message ID
        seen_ids = set()
        unique_messages = []
        for msg in all_messages:
            msg_id = msg.get("id")
            if msg_id and msg_id not in seen_ids:
                seen_ids.add(msg_id)
                unique_messages.append(msg)
        
        # Sort by date
        unique_messages.sort(key=lambda x: x.get("sentDateTime") or x.get("receivedDateTime") or "")
        
        return unique_messages
    
    def send_email(self, to_email: str, subject: str, body_html: str, user_email: Optional[str] = None):
        """
        Send an email via Microsoft Graph API.
        - For delegated auth: uses /me endpoint (works with Delegated Mail.Send permission)
        - For app-only auth: uses /users/{email} endpoint (requires Application Mail.Send permission)
        """
        # If using app-only auth (has client_secret), must specify user_email
        # If using delegated auth (no client_secret), can use /me endpoint
        if self.client_secret:
            # App-only authentication - requires Application permission
            if not user_email:
                raise Exception(
                    "For app-only authentication, user_email is required. "
                    "Also ensure you have Mail.Send as an APPLICATION permission (not Delegated)."
                )
            user_endpoint = f"/users/{user_email}"
        else:
            # Delegated authentication - can use /me endpoint (works with Delegated permission)
            user_endpoint = "/me"
        
        endpoint = f"{user_endpoint}/sendMail"
        
        message = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body_html
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            }
        }
        
        self.make_request("POST", endpoint, json=message)
    
    def create_draft_email(self, to_email: str, subject: str, body_html: str, user_email: Optional[str] = None) -> dict:
        """
        Create a draft email via Microsoft Graph API.
        - For delegated auth: uses /me endpoint (works with Delegated Mail.ReadWrite permission)
        - For app-only auth: uses /users/{email} endpoint (requires Application Mail.ReadWrite permission)
        """
        # If using app-only auth (has client_secret), must specify user_email
        # If using delegated auth (no client_secret), can use /me endpoint
        if self.client_secret:
            # App-only authentication - requires Application permission
            if not user_email:
                raise Exception(
                    "For app-only authentication, user_email is required. "
                    "Also ensure you have Mail.ReadWrite as an APPLICATION permission."
                )
            user_endpoint = f"/users/{user_email}"
        else:
            # Delegated authentication - can use /me endpoint (works with Delegated permission)
            user_endpoint = "/me"
        
        endpoint = f"{user_endpoint}/messages"
        
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body_html
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email
                    }
                }
            ]
        }
        
        return self.make_request("POST", endpoint, json=message)


class EmailAnalyzer:
    """Analyzes emails to determine if they require a reply - V2 with full conversation analysis."""
    
    # Latest GPT model
    MODEL = "gpt-4o"
    
    def __init__(self, openai_api_key: str):
        self.openai_client = openai.OpenAI(api_key=openai_api_key)
    
    def analyze_conversation_for_action(self, conversation_messages: List[dict], user_email: str, subject: str) -> Tuple[bool, str, str, str]:
        """
        MAIN METHOD: Analyze the ENTIRE conversation thread to determine if the user needs to take action.
        
        This method reads ALL messages in the conversation chronologically and determines:
        1. Is there an open question/request that hasn't been answered?
        2. Is the user (user_email) the one who needs to respond?
        3. Has someone already responded/closed the loop?
        
        Returns: (needs_action: bool, action_type: str, reason: str, confidence: str)
        - action_type: "user_reply_needed", "waiting_for_others", "closed", "no_action"
        - confidence: "high", "medium", or "low" (for edge cases)
        """
        if not conversation_messages:
            return False, "no_action", "No messages in conversation"
        
        # Build the full conversation thread for AI analysis
        conversation_text = self._build_conversation_thread(conversation_messages, user_email)
        
        # Get the user's name for the prompt
        user_name = user_email.split("@")[0].replace(".", " ").title()
        
        prompt = f"""You are analyzing an EMAIL CONVERSATION to determine if {user_name} ({user_email}) needs to take any action.

IMPORTANT RULES:
1. Read the ENTIRE conversation from oldest to newest
2. Check if any open questions/requests have been ANSWERED by subsequent messages
3. Determine if {user_name} specifically is being asked to respond, or if the question is directed at someone else
4. If the conversation is just forwarding info (FW:) with no question, no action is needed
5. If someone already responded to close the loop, no action is needed
6. If a meeting was scheduled and confirmed, no action is needed
7. If the question/request is directed at someone OTHER than {user_name} (e.g., "Aileen, please review" or "Nora, can you set up a call?"), then {user_name} does NOT need to take action
8. CRITICAL - WHO SENT THE LAST MESSAGE:
   - If {user_name} sent the LAST message (marked as "(YOU)"), action_type MUST be "waiting_for_others" or "no_action", NEVER "user_reply_needed"
   - When {user_name} sent the last message, the "reason" field MUST explicitly state WHO {user_name} is waiting for and WHAT response is expected
   - Example: "{user_name} sent comments on the proposal - waiting for John Smith to review and respond"
   - NEVER say "{user_name} needs to reply" if {user_name} sent the last message
9. CLOSURE DETECTION: If {user_name}'s last message contains closure signals like:
   - "Thank you", "Thanks", "Confirmed", "Access granted", "Done", "We will close this ticket"
   - "Sent you the keys", "Here are the credentials", "Access provided"
   - "No further action needed", "All set", "Looks good"
   - "Will reply to that other chain", "Resolved in another thread", "Handled in separate email"
   Then mark as "no_action" or "closed" - these are closure messages, not waiting for reply.
10. CREDENTIAL/INFO SHARING: If {user_name} sent credentials, API keys, DB credentials, passwords, or just informational content WITHOUT explicitly asking for confirmation/response, mark as "no_action". 
    These are one-way information sharing emails - no reply is expected unless explicitly requested.
    Examples: "AP Reminder tool DB credentials", "EPC RAG DB credentials", "Here are the API keys" (without "please confirm")
11. RESOLVED IN ANOTHER THREAD: If the conversation mentions being resolved/handled in another email/thread (e.g., "Will reply to that other chain", "Resolved in separate email", "Handled in another thread"), mark as "closed" or "no_action".
12. EDGE CASES: If you're uncertain whether action is needed (e.g., informational but might need acknowledgment),
    set confidence to "low" to flag for user review.

CONVERSATION (oldest first):
{conversation_text}

Answer in JSON format:
{{
    "needs_action": "Yes" or "No",
    "action_type": "user_reply_needed" or "waiting_for_others" or "closed" or "no_action",
    "reason": "<SPECIFIC and CONTEXTUAL explanation>",
    "directed_at": "<who is the request/question directed at, if applicable>",
    "confidence": "high" or "medium" or "low"
}}

Where:
- "user_reply_needed": {user_name} needs to respond to something
- "waiting_for_others": {user_name} asked something and is waiting for others to reply
- "closed": The conversation has been resolved/closed
- "no_action": No action needed (FYI, meeting scheduled, etc.)
- "confidence": 
  * "high": Clear action needed, no ambiguity
  * "medium": Action likely needed, but some uncertainty
  * "low": Edge case - might not need action, but flagging for review

In the "reason" field, provide a SPECIFIC and CONTEXTUAL explanation that naturally reflects the action_type:
- If action_type is "user_reply_needed": Explain what {user_name} needs to reply to and why
- If action_type is "waiting_for_others": Explain what {user_name} is waiting for from whom and why
- If action_type is "closed": Explain why the conversation is closed
- If action_type is "no_action": Explain why no action is needed

The reason should be:
- SPECIFIC: Reference actual content from the conversation (e.g., "John asked about the API status" not "needs attention")
- CONTEXTUAL: Explain the business/technical impact (e.g., "Workflow blocked - supplier approval cannot proceed" not "important")
- CLEAR: Help the user understand the exact situation

Examples of good reasons:
- (user_reply_needed) "John asked about the API status - {user_name} needs to provide an update on the deployment timeline"
- (waiting_for_others) "{user_name} identified a calculation error and sent corrections - waiting for Robert Schoenherr to review and confirm the fix"
- (waiting_for_others) "{user_name} asked about scheduling a call - waiting for John to provide his availability"
- (waiting_for_others) "{user_name} proposed a new structure with 10% pref - waiting for David to review and provide feedback"
- (closed) "Issue resolved - access granted and John confirmed receipt"
- (no_action) "Informational email - {user_name} shared DB credentials, no confirmation requested"

Avoid generic statements like "Needs attention", "Waiting for acknowledgment", or "Important email" without explaining why.
"""

        try:
            response = self.openai_client.chat.completions.create(
                model=self.MODEL,
                    messages=[
                    {"role": "system", "content": f"You are an expert email conversation analyzer. Your job is to determine if {user_name} ({user_email}) needs to take action on a conversation. Be STRICT - only flag conversations where {user_name} SPECIFICALLY needs to respond. If someone else is being asked, or if the loop has been closed, mark as no action needed. The reason you provide should naturally reflect the action_type you determine - be specific and contextual."},
                        {"role": "user", "content": prompt}
                    ],
                temperature=0.1,
                max_completion_tokens=500,  # Increased from 300 to allow complete, detailed explanations
                    response_format={"type": "json_object"}
                )
            
            result_text = response.choices[0].message.content.strip()
            result_json = json.loads(result_text)
            
            needs_action = result_json.get("needs_action", "").upper() == "YES"
            action_type = result_json.get("action_type", "no_action")
            reason = result_json.get("reason", "")
            directed_at = result_json.get("directed_at", "")
            confidence = result_json.get("confidence", "medium")  # Default to medium if not provided
            
            # Double-check: if directed at someone else, override to no action
            if directed_at and user_email.lower() not in directed_at.lower() and user_name.lower() not in directed_at.lower():
                if "everyone" not in directed_at.lower() and "all" not in directed_at.lower() and "team" not in directed_at.lower():
                    needs_action = False
                    action_type = "no_action"
                    reason = f"Request directed at {directed_at}, not {user_name}"
                    confidence = "high"  # High confidence when clearly directed at someone else
            
            return needs_action, action_type, reason, confidence
            
        except Exception as e:
            print(f"  Warning: AI conversation analysis failed: {e}")
            # Fallback to simple analysis
            return self._fallback_conversation_analysis(conversation_messages, user_email)
    
    def _build_conversation_thread(self, messages: List[dict], user_email: str) -> str:
        """Build a readable conversation thread from messages."""
        # Sort messages by date (oldest first)
        sorted_messages = sorted(messages, key=lambda x: x.get("sentDateTime") or x.get("receivedDateTime") or "")
        
        thread_parts = []
        user_email_lower = user_email.lower()
        
        for i, msg in enumerate(sorted_messages):
            # Get sender info
            from_email = msg.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
            from_name = msg.get("from", {}).get("emailAddress", {}).get("name", from_email)
            
            # Get recipients
            to_recipients = [r.get("emailAddress", {}).get("address", "") for r in msg.get("toRecipients", [])]
            to_str = ", ".join(to_recipients[:3])  # Limit to 3 recipients for brevity
            if len(to_recipients) > 3:
                to_str += f" +{len(to_recipients) - 3} more"
            
            # Get date
            date_str = msg.get("sentDateTime") or msg.get("receivedDateTime") or "Unknown date"
            if date_str and date_str != "Unknown date":
                try:
                    parsed_date = date_parser.parse(date_str)
                    date_str = parsed_date.strftime("%Y-%m-%d %H:%M")
                except:
                    pass
            
            # Get subject and body
            subject = msg.get("subject", "No subject")
            body = msg.get("body", {}).get("content", "") or msg.get("bodyPreview", "")
            body_text = self._extract_text_from_html(body)
            
            # Extract only new content (not quoted replies)
            new_content = self._extract_new_content(body_text)
            
            # Truncate if too long
            if len(new_content) > 500:
                new_content = new_content[:500] + "..."
            
            # Mark if this is from the user
            is_from_user = from_email.lower() == user_email_lower
            sender_label = f"{from_name} (YOU)" if is_from_user else from_name
            
            thread_parts.append(f"""
--- Message {i+1} ---
From: {sender_label}
To: {to_str}
Date: {date_str}
Subject: {subject}
Content: {new_content}
""")
        
        return "\n".join(thread_parts)
    
    def _extract_new_content(self, body_text: str) -> str:
        """Extract only new content from an email, excluding quoted replies."""
        if not body_text:
            return ""
        
        # Common patterns that indicate the start of quoted content
        quote_markers = [
            "From:",
            "-----Original Message-----",
            "On ",
            "> ",
            "wrote:",
            "Sent from",
            "________________________________",
            "-----Forwarded message-----",
            "Begin forwarded message:",
        ]
        
        lines = body_text.split('\n')
        new_content_lines = []
        
        for line in lines:
            # Check if this line starts quoted content
            is_quote_start = False
            for marker in quote_markers:
                if marker.lower() in line.lower()[:50]:  # Check first 50 chars
                    is_quote_start = True
                    break
            
            if is_quote_start:
                break
            
            new_content_lines.append(line)
        
        return '\n'.join(new_content_lines).strip()
    
    def _fallback_conversation_analysis(self, messages: List[dict], user_email: str) -> Tuple[bool, str, str]:
        """Fallback analysis if AI fails."""
        if not messages:
            return False, "no_action", "No messages"
        
        # Get latest message
        sorted_msgs = sorted(messages, key=lambda x: x.get("sentDateTime") or x.get("receivedDateTime") or "", reverse=True)
        latest = sorted_msgs[0]
        
        latest_from = latest.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        is_from_user = latest_from == user_email.lower()
        
        if is_from_user:
            return True, "waiting_for_others", "You sent the last message"
        else:
            return True, "user_reply_needed", "Someone else sent the last message"
    
    def requires_reply(self, subject: str, body: str) -> Tuple[bool, str]:
        """
        LEGACY METHOD: Analyze a single email to determine if it requires a reply.
        For backward compatibility. Prefer analyze_conversation_for_action() for full thread analysis.
        """
        body_text = self._extract_text_from_html(body) if body else ""
        body_preview = body_text[:1500] if len(body_text) > 1500 else body_text
        
        excerpt = f"Subject: {subject}\n\nBody: {body_preview}"
        
        prompt = f"""Analyze this email and determine if it expects a reply.

Answer ONLY in JSON:
{{"reply_expected": "Yes" or "No", "reason": "<short explanation>"}}

Rules:
- Mark "No" if this email is a confirmation, update, acknowledgement, FYI, or provides an attachment
- Mark "No" if this email completes a prior request ("I've scheduled...", "Please find attached...", "Here you go", "Thanks", "Sounds good")
- Mark "No" if this is a forwarded email (FW:) with no new question
- Mark "Yes" ONLY if the email asks a NEW question or requests action that hasn't been addressed

EMAIL:
\"\"\"{excerpt}\"\"\"
"""

        try:
            response = self.openai_client.chat.completions.create(
                model=self.MODEL,
                messages=[
                    {"role": "system", "content": "You are a strict email analysis assistant. Be conservative - only mark 'Yes' if there's a clear NEW question or action request. Mark 'No' for confirmations, updates, acknowledgements, forwards, or messages that complete prior requests."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_completion_tokens=150,
                response_format={"type": "json_object"}
            )
            
            result_text = response.choices[0].message.content.strip()
            result_json = json.loads(result_text)
            
            reply_expected = result_json.get("reply_expected", "").upper() == "YES"
            reason = result_json.get("reason", "")
            
            return reply_expected, reason
            
        except Exception as e:
            print(f"  Warning: AI analysis failed: {e}")
            return False, "Analysis failed - defaulting to no reply needed"
    
    def analyze_urgency(self, subject: str, body: str) -> Tuple[bool, str]:
        """
        Use AI to determine if an email is urgent/time-sensitive.
        Returns: (is_urgent: bool, reason: str)
        """
        body_text = self._extract_text_from_html(body) if body else ""
        body_preview = body_text[:800] if len(body_text) > 800 else body_text
        
        excerpt = f"Subject: {subject}\n\nBody: {body_preview}"
        
        prompt = f"""Analyze this email and determine if it is URGENT or TIME-SENSITIVE.

Answer ONLY in JSON format:
{{"is_urgent": "Yes" or "No", "reason": "<detailed explanation of why it is or isn't urgent>"}}

An email is URGENT if it:
- Has a deadline within the next 2-3 days that hasn't passed
- Requires immediate action or response
- Uses words like urgent, ASAP, critical, immediately in the NEW content (not quoted replies)
- Has time-sensitive consequences if not addressed quickly
- The urgency is CONTEXTUAL - e.g., "system down", "blocking issue", "deadline tomorrow"

An email is NOT urgent if it:
- Is just informational or FYI
- Is a calendar invite acceptance/decline
- Has no specific deadline
- The urgent keyword appears only in quoted/forwarded content
- Is a routine update
- The urgency is historical (e.g., "we needed this yesterday" in a closed thread)
- The email is a closure/confirmation message (e.g., "Thanks", "Confirmed", "Done")

CONTEXT-AWARE ANALYSIS:
- Check if urgency keywords are in NEW content vs quoted/forwarded content
- Consider the context: Is this an active issue or a historical reference?
- If the email is a reply that resolves/closes the issue, it's NOT urgent
- If the email is informational without a deadline, it's NOT urgent

IMPORTANT: Provide a COMPLETE and DETAILED explanation in the "reason" field. Explain:
- What makes it urgent (deadline, keywords, context) - be SPECIFIC
- Why immediate action is needed - the ACTUAL reason, not generic statements
- What the consequences might be if not addressed - be specific about impact
- If NOT urgent, explain why (e.g., "informational only", "deadline already passed", "urgency only in quoted content")

EMAIL:
\"\"\"{excerpt}\"\"\"
"""

        try:
            response = self.openai_client.chat.completions.create(
                model=self.MODEL,
                messages=[
                    {"role": "system", "content": "You are an email urgency analyzer. Be conservative - only mark truly urgent items where the urgency is in the NEW content, not quoted replies. Always provide detailed, complete explanations. Be context-aware - consider if urgency is historical or active."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_completion_tokens=300,  # Increased from 100 to allow complete explanations
                response_format={"type": "json_object"}
            )
            
            result_text = response.choices[0].message.content.strip()
            result_json = json.loads(result_text)
            
            is_urgent = result_json.get("is_urgent", "").upper() == "YES"
            reason = result_json.get("reason", "")
            
            return is_urgent, reason
            
        except Exception as e:
            print(f"  ❌ Error analyzing urgency: {e}")
            return False, ""
    
    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content."""
        if not html_content:
            return ""
        
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', html_content)
        # Decode HTML entities
        text = html.unescape(text)
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        return text


def extract_to_recipients(message: dict) -> List[str]:
    """Extract 'To' recipient email addresses from a message."""
    recipients = []
    to_recipients = message.get("toRecipients", [])
    for recipient in to_recipients:
        email = recipient.get("emailAddress", {}).get("address", "")
        if email:
            recipients.append(email.lower())
    return recipients


def extract_recipient_display(message: dict, recipient_type: str = "toRecipients") -> List[str]:
    """
    Extract recipient display names/emails for display purposes.
    For shared inboxes or DLs, returns the shared email address or DL name.
    """
    recipients = []
    recip_list = message.get(recipient_type, [])
    for recipient in recip_list:
        email_addr = recipient.get("emailAddress", {})
        name = email_addr.get("name", "")
        email = email_addr.get("address", "")
        # For shared inboxes/DLs, prefer the email address (which is the shared email)
        # If name is available and different, use "Name <email>", otherwise just email
        if email:
            if name and name.lower() != email.lower():
                display = f"{name} ({email})"
            else:
                display = email
            recipients.append(display)
    return recipients


def _clean_subject(subject: str) -> str:
    """Remove RE:, FW:, FWD: prefixes from subject for comparison."""
    clean = subject.strip()
    prefixes = ["RE:", "Re:", "re:", "FW:", "Fw:", "fw:", "FWD:", "Fwd:", "fwd:"]
    changed = True
    while changed:
        changed = False
        for prefix in prefixes:
            if clean.startswith(prefix):
                clean = clean[len(prefix):].strip()
                changed = True
    return clean.lower()


def extract_all_recipients(message: dict) -> List[str]:
    """Extract all recipient email addresses (To, CC, BCC) from a message."""
    all_recipients = []
    
    # Extract To recipients
    to_recipients = extract_to_recipients(message)
    all_recipients.extend(to_recipients)
    
    # Extract CC recipients
    cc_recipients = message.get("ccRecipients", [])
    for recipient in cc_recipients:
        email = recipient.get("emailAddress", {}).get("address", "").lower()
        if email and email not in all_recipients:
            all_recipients.append(email)
    
    # Extract BCC recipients
    bcc_recipients = message.get("bccRecipients", [])
    for recipient in bcc_recipients:
        email = recipient.get("emailAddress", {}).get("address", "").lower()
        if email and email not in all_recipients:
            all_recipients.append(email)
    
    return all_recipients


def format_datetime_et(dt_obj: dt.datetime) -> str:
    """Convert datetime to ET timezone and format as string."""
    if not dt_obj:
        return "N/A"
    
    # ET timezone (handles EST/EDT automatically)
    et_tz = tz.gettz("America/New_York")
    
    # Convert to ET if timezone-aware, otherwise assume UTC
    if dt_obj.tzinfo:
        et_dt = dt_obj.astimezone(et_tz)
    else:
        et_dt = dt_obj.replace(tzinfo=tz.UTC).astimezone(et_tz)
    
    # Format as "Dec 22, 2025 2:30 PM ET"
    return et_dt.strftime("%b %d, %Y %I:%M %p ET")


def get_message_datetime(message: dict) -> dt.datetime:
    """
    Extract the best available datetime from a message.
    Prefers sentDateTime for sent messages, receivedDateTime for received messages.
    Returns a timezone-aware datetime.
    """
    # Try sentDateTime first (for sent messages)
    sent_dt_str = message.get("sentDateTime")
    if sent_dt_str:
        try:
            parsed = date_parser.parse(sent_dt_str)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=tz.UTC)
            return parsed
        except:
            pass
    
    # Try receivedDateTime (for received messages)
    recv_dt_str = message.get("receivedDateTime")
    if recv_dt_str:
        try:
            parsed = date_parser.parse(recv_dt_str)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=tz.UTC)
            return parsed
        except:
            pass
    
    # Fallback to a very old date (should rarely happen)
    return dt.datetime(1970, 1, 1, tzinfo=tz.UTC)


def get_latest_message_in_conversation(conversation_messages: List[dict], user_email: str) -> Optional[dict]:
    """
    Get the latest message in a conversation thread.
    Returns the message dict if found, None otherwise.
    """
    if not conversation_messages:
        return None
    
    # Filter out messages with invalid dates
    valid_messages = []
    for msg in conversation_messages:
        msg_date = get_message_datetime(msg)
        # Skip messages from 1970 (fallback date) unless it's the only message
        if msg_date.year > 1970 or len(conversation_messages) == 1:
            valid_messages.append((msg_date, msg))
    
    if not valid_messages:
        # If no valid messages, return the first message as fallback
        return conversation_messages[0]
    
    # Sort by date (latest first) and return the latest message
    valid_messages.sort(key=lambda x: x[0], reverse=True)
    return valid_messages[0][1]


def is_message_from_user(message: dict, user_email: str) -> bool:
    """
    Check if a message is sent by the user.
    Uses multiple methods to reliably identify user's messages:
    1. Check 'from' field (most reliable for sent messages)
    2. Check 'sender' field
    3. Check if message is in Sent Items folder
    4. Check if message only has sentDateTime (no receivedDateTime) - indicates sent message
    """
    if not message:
        return False
    
    user_email_lower = user_email.lower() if user_email else ""
    
    # Method 1: Check 'from' field (most reliable for sent messages)
    from_field = message.get("from", {}).get("emailAddress", {})
    from_email = from_field.get("address", "").lower()
    if from_email and from_email == user_email_lower:
        return True
    
    # Method 2: Check 'sender' field
    sender = message.get("sender", {}).get("emailAddress", {})
    sender_email = sender.get("address", "").lower()
    if sender_email and sender_email == user_email_lower:
        return True
    
    # Method 3: Check if message is from Sent Items folder
    # Sent Items folder ID typically contains "sentitems" (case-insensitive)
    parent_folder_id = message.get("parentFolderId", "").lower()
    if "sentitems" in parent_folder_id:
        return True
    
    # Method 4: Check if message has sentDateTime but no receivedDateTime
    # Sent messages have sentDateTime, received messages have receivedDateTime
    has_sent_date = bool(message.get("sentDateTime"))
    has_received_date = bool(message.get("receivedDateTime"))
    
    # If it has sentDateTime but no receivedDateTime, it's likely a sent message
    # Also check that we don't have a sender field pointing to someone else
    if has_sent_date and not has_received_date:
        # Make sure sender (if exists) isn't someone else
        if not sender_email or sender_email == user_email_lower or not sender_email.strip():
            return True
    
    return False


def extract_new_content_from_email(body: str) -> str:
    """
    Extract only the NEW content from an email, excluding quoted replies.
    This helps avoid detecting keywords from old quoted messages.
    """
    if not body:
        return ""
    
    # Common patterns that indicate the start of quoted content
    quote_markers = [
        "From:",  # Outlook-style quote
        "-----Original Message-----",
        "On ",  # Gmail-style "On <date>, <person> wrote:"
        "> ",  # Traditional quote marker
        "wrote:",
        "Sent from",
        "________________________________",  # Outlook separator
        "-----Forwarded message-----",
    ]
    
    lines = body.split('\n')
    new_content_lines = []
    
    for line in lines:
        # Check if this line starts quoted content
        is_quote_start = False
        for marker in quote_markers:
            if marker.lower() in line.lower():
                is_quote_start = True
                break
        
        if is_quote_start:
            # Stop collecting - rest is quoted content
            break
        
        new_content_lines.append(line)
    
    return '\n'.join(new_content_lines)


def detect_priority_keywords(subject: str, body: str) -> Tuple[bool, List[str]]:
    """
    Detect if an email is high priority based on keywords.
    Only checks the NEW content of the email, not quoted replies.
    Returns: (has_keywords: bool, keywords_found: List[str])
    """
    # Extract only new content, excluding quoted replies
    new_body = extract_new_content_from_email(body)
    
    # Check subject (always check full subject) + new body content only
    text = f"{subject} {new_body}".lower()
    keywords_found = []
    
    # Check urgent keywords
    for keyword in URGENT_KEYWORDS:
        if keyword.lower() in text:
            keywords_found.append(keyword.upper())
    
    # Check deadline keywords
    for keyword in DEADLINE_KEYWORDS:
        if keyword.lower() in text:
            keywords_found.append(keyword.upper())
    
    # Check action keywords
    for keyword in ACTION_KEYWORDS:
        if keyword.lower() in text:
            keywords_found.append(keyword.upper())
    
    # Remove duplicates while preserving order
    seen = set()
    unique_keywords = []
    for kw in keywords_found:
        if kw not in seen:
            seen.add(kw)
            unique_keywords.append(kw)
    
    has_keywords = len(unique_keywords) > 0
    return has_keywords, unique_keywords[:3]  # Return max 3 keywords for display


def check_reply_received(sent_message: dict, conversation_messages: List[dict], 
                         all_recipients: List[str], my_email: str, 
                         deadline: dt.datetime) -> Tuple[bool, Optional[dt.datetime], Optional[str]]:
    """
    Check if any recipient (To, CC, or BCC) replied within the deadline.
    Also checks if the conversation appears to be closed (recent activity suggests it's resolved).
    Returns: (has_reply, last_reply_date, last_reply_sender)
    """
    sent_date = date_parser.parse(sent_message["sentDateTime"])
    if sent_date.tzinfo is None:
        sent_date = sent_date.replace(tzinfo=tz.UTC)
    
    my_email_lower = my_email.lower() if my_email else ""
    latest_reply_date = None
    latest_reply_sender = None
    
    # Track all replies (not just within deadline) to check if conversation is closed
    all_replies = []
    
    for msg in conversation_messages:
        # Skip the original sent message
        if msg.get("id") == sent_message.get("id"):
            continue
        
        # Check if this is a received message (has receivedDateTime)
        received_date = msg.get("receivedDateTime")
        if not received_date:
            continue
        
        received_dt = date_parser.parse(received_date)
        if received_dt.tzinfo is None:
            received_dt = received_dt.replace(tzinfo=tz.UTC)
        
        # Must be after sent date
        if received_dt <= sent_date:
            continue
        
        # Get sender email
        sender = msg.get("sender", {}).get("emailAddress", {})
        sender_email = sender.get("address", "").lower()
        
        # Skip if sender is me
        if sender_email == my_email_lower:
            continue
        
        # Check if sender is one of the recipients (To, CC, or BCC)
        if sender_email in all_recipients:
            # This is a valid reply from a recipient
            all_replies.append((received_dt, sender_email, sender.get("name", "")))
            
            # Check if within deadline
            if received_dt <= deadline:
                if latest_reply_date is None or received_dt > latest_reply_date:
                    latest_reply_date = received_dt
                    sender_name = sender.get("name", "")
                    if sender_name and sender_name.lower() != sender_email.lower():
                        latest_reply_sender = f"{sender_name} ({sender_email})"
                    else:
                        latest_reply_sender = sender_email
    
    # Check if conversation appears closed (has recent replies, even if after deadline)
    # If there are replies after the deadline, the conversation might be active/closed
    has_reply = latest_reply_date is not None
    
    # If no reply within deadline but there are replies after deadline, get the latest
    if not has_reply and all_replies:
        # Sort by date and get the latest
        all_replies.sort(key=lambda x: x[0], reverse=True)
        latest_reply_dt, latest_sender_email, latest_sender_name = all_replies[0]
        latest_reply_date = latest_reply_dt
        if latest_sender_name and latest_sender_name.lower() != latest_sender_email.lower():
            latest_reply_sender = f"{latest_sender_name} ({latest_sender_email})"
        else:
            latest_reply_sender = latest_sender_email
    
    return has_reply, latest_reply_date, latest_reply_sender


def create_outlook_link(web_link: str, message_id: str) -> str:
    """
    Create a link to open the email in Outlook.
    Uses the webLink from Graph API which opens in Outlook web/desktop app.
    """
    # Use the webLink directly - it should open in Outlook web or desktop app
    # The webLink from Microsoft Graph API is the official link
    if web_link:
        # Ensure it's a valid URL
        if web_link.startswith("http://") or web_link.startswith("https://"):
            return web_link
        # If it's a relative path, make it absolute
        if web_link.startswith("/"):
            return f"https://outlook.office.com{web_link}"
    
    # Fallback: construct a web link using message ID
    if message_id:
        # Try Outlook web link format
        return f"https://outlook.office.com/mail/item/{message_id}"
    
    # Last resort: return empty string (will show as "N/A" in table)
    return ""


def build_section_table(entries: List[Dict], section_type: str) -> str:
    """
    Build HTML table for a specific section.
    section_type: 'urgent', 'recent_important', 'hanging', or 'auto_closed'
    """
    if not entries:
        return ""
    
    # Different headers based on section type - WIDER COLUMNS for text-heavy content - STANDARDIZED FONT SIZE
    # Only auto_closed section has an extra "Reason" column
    if section_type == 'auto_closed':
        extra_col = '<th style="border:1px solid #ddd; padding:12px; text-align:left; font-weight:bold; font-size:13px; width:25%;">Reason</th>'
    else:
        extra_col = ''
    
    header = f"""
    <table style="border-collapse:collapse; width:100%; font-family:Segoe UI, Arial, sans-serif; margin:10px 0;">
    <thead>
        <tr style="background-color:#f2f2f2;">
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:2%;">#</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:14%;">Subject</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:8%;">Last From</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:4%;">Age</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:10%;">Last Message (ET)</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:20%;">Conversation Summary</th>
        {extra_col}
        <th style="border:1px solid #ddd; padding:10px; text-align:left; font-weight:bold; font-size:13px; width:7%;">Status</th>
        <th style="border:1px solid #ddd; padding:10px; text-align:center; font-weight:bold; font-size:13px; width:16%;">Action</th>
        </tr>
    </thead>
    <tbody>
    """
    
    rows_html = ""
    for idx, e in enumerate(entries, 1):
        subject_raw = e.get("subject", "") or ""
        # NO TRUNCATION - show full subject
        subject_escaped = html.escape(subject_raw)
        
        # Get web link for clickable subject
        web_link = e.get("web_link", "")
        
        # Visual indicator for low confidence (edge cases) - NO section icons on individual rows
        confidence = e.get("confidence", "medium")
        
        # Add confidence indicator only (no section icons on rows) - STANDARDIZED FONT SIZE
        # Make subject clickable if web_link exists
        if confidence == "low":
            # Add a subtle indicator for edge cases
            if web_link:
                subject = f'<a href="{html.escape(web_link)}" style="color:#0066cc; text-decoration:none; font-size:13px; opacity:0.85; border-left:3px solid #ff9800; padding-left:5px; display:inline-block;">⚠️ {subject_escaped}</a>'
            else:
                subject = f'<span style="font-size:13px; opacity:0.85; border-left:3px solid #ff9800; padding-left:5px; display:inline-block;">⚠️ {subject_escaped}</span>'
        else:
            # No icon, just the subject with consistent font size
            if web_link:
                subject = f'<a href="{html.escape(web_link)}" style="color:#0066cc; text-decoration:none; font-size:13px;">{subject_escaped}</a>'
            else:
                subject = f'<span style="font-size:13px;">{subject_escaped}</span>'
        
        # Last message from - format as proper name - STANDARDIZED FONT SIZE
        last_msg_from = e.get("last_msg_from", "") or "N/A"
        last_msg_from = format_sender_name(last_msg_from)
        last_msg_from_escaped = html.escape(last_msg_from)[:25]
        last_msg_from_html = f'<span style="font-size:13px;">{last_msg_from_escaped}</span>'
        
        # Age in days (or days inactive for auto_closed) - STANDARDIZED FONT SIZE
        if section_type == 'auto_closed':
            days_inactive = e.get("days_inactive", 0)
            age_str = f"{days_inactive:.0f}d"
        else:
            days_old = e.get("days_old", 0)
            age_str = f"{days_old:.1f}d"
        age_str_html = f'<span style="font-size:13px;">{age_str}</span>'
        
        # Conversation Summary (from ai_reason) - NO TRUNCATION, STANDARDIZED FONT SIZE
        conversation_summary = e.get("ai_reason", "") or "No summary available"
        conversation_summary_html = f'<span style="font-size:13px; line-height:1.4; color:#333;">{html.escape(conversation_summary)}</span>'
        
        # Last Message Date/Time in ET - STANDARDIZED FONT SIZE
        last_msg_date = e.get("last_msg_date")
        last_msg_et = format_datetime_et(last_msg_date) if last_msg_date else "N/A"
        last_msg_et_html = f'<span style="font-size:13px; color:#666;">{html.escape(last_msg_et)}</span>'
        
        # Extra column content - only for auto_closed section
        if section_type == 'auto_closed':
            # Show reason for auto-closure
            reason = e.get("reason", "") or "Inactive for 14+ days"
            extra_content = f'<span style="font-size:13px; line-height:1.4; color:#666;">{html.escape(reason)}</span>'
            extra_cell = f'<td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{extra_content}</td>'
        else:
            extra_cell = ''
        
        # Action needed (with color coding) - STANDARDIZED FONT SIZE (12px for badges)
        if section_type == 'auto_closed':
            action_needed = "Auto-Closed"
            action_style = "background-color:#f5f5f5; color:#757575; font-weight:bold; padding:4px 8px; border-radius:4px; font-size:12px;"
        else:
            action_needed = e.get("action_needed", "")
            if action_needed == "You need to reply":
                action_style = "background-color:#ffebee; color:#c62828; font-weight:bold; padding:4px 8px; border-radius:4px; font-size:12px;"
            else:
                action_style = "background-color:#e3f2fd; color:#1565c0; font-weight:bold; padding:4px 8px; border-radius:4px; font-size:12px;"
        action_html = f'<span style="{action_style}">{html.escape(action_needed)}</span>'
        
        # "Mark as Dealt With" button - Simple table-based, green background, clickable
        conv_id = e.get("conversation_id", "")
        msg_id = e.get("latest_message_id", "")
        email = e.get("user_email", "")
        subject_for_url = e.get("subject", "")

        if conv_id and msg_id and email:
            # URL-encode the subject to handle special characters
            encoded_subject = quote(subject_for_url, safe='')
            mark_dealt_url = f"{html.escape(WEBHOOK_API_URL)}/api/mark-dealt-with?conversationId={html.escape(conv_id)}&latestMessageId={html.escape(msg_id)}&userEmail={html.escape(email)}&subject={encoded_subject}"
            # Simple table button - green background, white text, fully clickable
            mark_dealt_html = f'''<table border="0" cellspacing="0" cellpadding="0"><tr><td bgcolor="#4CAF50" style="padding:8px 12px;"><a href="{mark_dealt_url}" target="_blank" style="color:#ffffff;font-family:Segoe UI,Arial,sans-serif;font-size:11px;font-weight:bold;text-decoration:none;white-space:nowrap;">✓ Mark as Dealt With</a></td></tr></table>'''
        else:
            mark_dealt_html = '<span style="font-size:11px; color:#999;">N/A</span>'
        
        rows_html += f"""
        <tr>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top; text-align:center; font-size:13px;">{idx}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{subject}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{last_msg_from_html}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top; text-align:center;">{age_str_html}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{last_msg_et_html}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{conversation_summary_html}</td>
        {extra_cell}
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top;">{action_html}</td>
        <td style="border:1px solid #ddd; padding:8px; vertical-align:top; text-align:center;">{mark_dealt_html}</td>
        </tr>
        """
    
    footer = """
      </tbody>
    </table>
    """
    
    return header + rows_html + footer


def format_sender_name(email_or_name: str) -> str:
    """
    Format sender display name nicely.
    - If it's an email like "Vikas.Agrawal@desri.com", extract and format as "Vikas Agrawal"
    - If it already has a name, use it
    """
    if not email_or_name:
        return "Unknown"
    
    # If it's an email address, extract the name part
    if "@" in email_or_name:
        name_part = email_or_name.split("@")[0]
        # Replace dots and underscores with spaces, then title case
        name_formatted = name_part.replace(".", " ").replace("_", " ").replace("-", " ")
        # Title case each word
        return " ".join(word.capitalize() for word in name_formatted.split())
    
    return email_or_name


def build_enhanced_digest(urgent_emails: List[Dict], recent_important: List[Dict], hanging_emails: List[Dict], auto_closed_emails: List[Dict], stats: Dict, user_name: str = "") -> str:
    """
    Build the enhanced HTML digest with 4 sections:
    1. Urgent - Action Required (has priority keywords or AI detected urgency)
    2. Recent but Important (< 2 days but needs attention)
    3. Hanging Conversations (2+ days, no reply)
    4. Auto-Closed - Inactive Conversations (14+ days inactive)
    """
    
    total_processed = stats.get('total_processed', 0)
    total_attention = len(urgent_emails) + len(recent_important) + len(hanging_emails)
    
    # Header with analysis summary - Outlook-friendly design with dark text - WIDER TEMPLATE
    digest_html = f"""
    <div style="font-family:Segoe UI, Arial, sans-serif; max-width:1400px; margin:0 auto; padding:20px;">
      <div style="border-bottom:3px solid #1a237e; padding:15px 0;">
        <h1 style="margin:0; font-size:24px; color:#1a237e;">📧 Email Follow-up Daily Digest{f' - {user_name}' if user_name else ''}</h1>
        <p style="margin:5px 0 0 0; font-size:13px; color:#666666;">Generated: {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
      </div>
      
      <div style="background:#e8eaf6; padding:15px; border-left:1px solid #ddd; border-right:1px solid #ddd;">
        <p style="margin:0; font-size:13px; color:#333;">
          📊 <strong>Analysis Summary:</strong> Analyzed <strong>{total_processed}</strong> sent emails from the last <strong>{LOOKBACK_DAYS} days</strong>. 
          Found <strong>{total_attention}</strong> email(s) requiring your attention. Below is the detailed breakdown.
        </p>
      </div>
      
      <div style="background:#f5f5f5; padding:15px; border-left:1px solid #ddd; border-right:1px solid #ddd;">
        <table style="width:100%;">
          <tr>
            <td style="text-align:center; padding:10px;">
              <div style="font-size:28px; font-weight:bold; color:#c62828;">{len(urgent_emails)}</div>
              <div style="font-size:13px; color:#666;">🔴 Urgent</div>
            </td>
            <td style="text-align:center; padding:10px;">
              <div style="font-size:28px; font-weight:bold; color:#ef6c00;">{len(recent_important)}</div>
              <div style="font-size:13px; color:#666;">🟠 Recent Important</div>
            </td>
            <td style="text-align:center; padding:10px;">
              <div style="font-size:28px; font-weight:bold; color:#1565c0;">{len(hanging_emails)}</div>
              <div style="font-size:13px; color:#666;">⏳ Hanging</div>
            </td>
            <td style="text-align:center; padding:10px;">
              <div style="font-size:28px; font-weight:bold; color:#2e7d32;">{stats.get('no_action', 0)}</div>
              <div style="font-size:13px; color:#666;">✅ No Action</div>
            </td>
          </tr>
        </table>
      </div>
      
    """
    
    # Section 1: Urgent
    if urgent_emails:
        digest_html += f"""
      <div style="border:1px solid #ddd; border-top:none; padding:15px;">
        <h2 style="color:#c62828; margin:0 0 10px 0; font-size:18px; border-bottom:2px solid #c62828; padding-bottom:8px;">
          🔴 URGENT - Action Required - {len(urgent_emails)}
        </h2>
        <p style="color:#666; font-size:13px; margin:0 0 10px 0;">High-priority emails detected via AI analysis (AI understands urgency from context)</p>
        {build_section_table(urgent_emails, 'urgent')}
      </div>
        """
    
    # Section 2: Recent but Important
    if recent_important:
        digest_html += f"""
      <div style="border:1px solid #ddd; border-top:none; padding:15px;">
        <h2 style="color:#ef6c00; margin:0 0 10px 0; font-size:18px; border-bottom:2px solid #ef6c00; padding-bottom:8px;">
          🟠 Recent but Important - {len(recent_important)}
        </h2>
        <p style="color:#666; font-size:13px; margin:0 0 10px 0;">Less than {RECENT_THRESHOLD_DAYS} days old, but AI detected these need attention</p>
        {build_section_table(recent_important, 'recent_important')}
      </div>
        """
    
    # Section 3: Hanging Conversations (WITH ⚠️ EXPLANATION)
    if hanging_emails:
        digest_html += f"""
      <div style="border:1px solid #ddd; border-top:none; padding:15px;">
        <h2 style="color:#1565c0; margin:0 0 10px 0; font-size:18px; border-bottom:2px solid #1565c0; padding-bottom:8px;">
          ⏳ Hanging Conversations - {len(hanging_emails)}
        </h2>
        <p style="color:#666; font-size:13px; margin:0 0 10px 0;">Conversations waiting {REPLY_WAIT_DAYS}+ days for a response</p>
        <p style="color:#856404; font-size:13px; margin:0 0 10px 0; padding:8px; background-color:#fff3cd; border-left:3px solid #ff9800; border-radius:3px;">
          <strong>⚠️ Note:</strong> Emails marked with <strong>⚠️</strong> indicate low confidence / edge cases that may need manual review.
        </p>
        {build_section_table(hanging_emails, 'hanging')}
      </div>
        """
    
    # Section 4: Auto-Closed - Inactive Conversations
    if auto_closed_emails:
        digest_html += f"""
      <div style="border:1px solid #ddd; border-top:none; padding:15px;">
        <h2 style="color:#9e9e9e; margin:0 0 10px 0; font-size:18px; border-bottom:2px solid #9e9e9e; padding-bottom:8px;">
          ⚪ Auto-Closed - Inactive Conversations - {len(auto_closed_emails)}
        </h2>
        <p style="color:#666; font-size:13px; margin:0 0 10px 0;">
          These conversations were inactive for 14+ days and were automatically marked as closed. 
          Review to confirm closure or reopen if needed.
        </p>
        {build_section_table(auto_closed_emails, 'auto_closed')}
      </div>
        """
    
    # No items message
    if not urgent_emails and not recent_important and not hanging_emails:
        digest_html += """
      <div style="border:1px solid #ddd; border-top:none; padding:40px; text-align:center;">
        <div style="font-size:48px;">🎉</div>
        <h2 style="color:#2e7d32; margin:10px 0;">All caught up!</h2>
        <p style="color:#666;">No emails requiring immediate attention.</p>
      </div>
        """
    
    # Close main container
    digest_html += """
    </div>
    """
    
    return digest_html


def analyze_user_mailbox(user_email: str, graph_client: GraphAPIClient, analyzer: EmailAnalyzer) -> Tuple[List[Dict], List[Dict], List[Dict], List[Dict], Dict]:
    """
    V2: Analyze a single user's mailbox using FULL CONVERSATION analysis.
    
    Key improvements:
    1. AI reads the ENTIRE conversation thread, not just the latest message
    2. AI checks if the user is specifically being asked to respond
    3. AI detects if someone else already responded/closed the loop
    4. Better forward detection
    5. Uses GPT-4o for more accurate analysis
    6. Time-based auto-closure for inactive conversations (14+ days)
    7. Detection of instructions/info sent by user
    8. Multi-thread awareness with contextual notes
    
    Returns: (urgent_emails, recent_important, hanging_emails, auto_closed_emails, stats_dict)
    """
    from datetime import datetime
    
    # Calculate date range
    now = dt.datetime.now(tz=tz.UTC)
    lookback_date = now - dt.timedelta(days=LOOKBACK_DAYS)
    
    print(f"\n{'='*80}")
    print(f"Analyzing mailbox: {user_email} (V2 - Full Conversation Analysis)")
    print(f"Using model: {analyzer.MODEL}")
    print(f"{'='*80}")
    print(f"Analyzing sent emails from the past {LOOKBACK_DAYS} days...")
    print(f"AI will read ENTIRE conversation threads to determine action needed...")
    
    # Get sent messages
    sent_messages = graph_client.get_sent_messages(user_email, lookback_date)
    print(f"Found {len(sent_messages)} sent messages")
    
    # Apply test limit if set
    if TEST_LIMIT:
        sent_messages = sent_messages[:TEST_LIMIT]
        print(f"⚠️ TEST MODE: Limiting to first {TEST_LIMIT} emails")
    
    # Process messages - categorize into 4 groups
    urgent_emails = []        # Has priority keywords AND needs action
    recent_important = []     # < 2 days AND needs action
    hanging_emails = []       # 2+ days AND needs action
    auto_closed_emails = []   # Conversations auto-closed due to inactivity (14+ days)
    skipped_count = 0
    
    conversation_cache = {}
    # Track processed conversations: conversationId -> (first_subject, first_msg_index, result)
    processed_conversations = {}
    
    user_email_lower = user_email.lower() if user_email else ""
    user_name = user_email.split("@")[0].replace(".", " ").title()
    
    for idx, sent_msg in enumerate(sent_messages, 1):
        subject = sent_msg.get('subject', 'No subject')
        
        try:
            print(f"\nProcessing message {idx}/{len(sent_messages)}: {subject[:50]}")
            
            # Extract all recipients (To, CC, BCC)
            all_recipients = extract_all_recipients(sent_msg)
            
            # Basic validations
            if not all_recipients:
                skipped_count += 1
                print(f"  Skipping: No recipients found")
                continue
            
            # Skip emails sent to self (if all recipients are the user)
            if all(recipient.lower() == user_email_lower for recipient in all_recipients):
                skipped_count += 1
                print(f"  Skipping: Email sent to self")
                continue
            
            # Skip no-reply addresses
            if any("noreply" in email or "no-reply" in email or "donotreply" in email for email in all_recipients):
                skipped_count += 1
                print(f"  Skipping: No-reply address")
                continue
            
            # Get conversationId
            conversation_id = sent_msg.get("conversationId")
            if not conversation_id:
                skipped_count += 1
                print(f"  Warning: No conversationId found, skipping...")
                continue
            
            # Check if this is a FORWARD from the user (FW:, Fw:, FWD:)
            is_forward = any(subject.upper().startswith(prefix) for prefix in ["FW:", "FWD:"])
            
            # Skip if we've already processed this conversation
            # EXCEPTION: If current message is a forward, it might deserve separate analysis
            # because the user might have forwarded to different people for different reasons
            if conversation_id in processed_conversations:
                orig_info = processed_conversations[conversation_id]
                original_subject = orig_info['subject']
                original_msg_idx = orig_info['msg_idx']
                original_result = orig_info.get('result', 'unknown')
                
                # If both have same base subject (ignoring RE:/FW:), skip as duplicate
                clean_current = _clean_subject(subject)
                clean_original = _clean_subject(original_subject)
                
                if clean_current == clean_original:
                    print(f"  Skipping: Already analyzed in msg #{original_msg_idx}")
                    print(f"    Original: '{original_subject[:50]}' -> {original_result}")
                    skipped_count += 1
                    continue
                else:
                    # Different subjects in same conversation - might be a branch
                    # Check if this forward should be analyzed separately
                    if is_forward:
                        print(f"  Note: Forward in same conversation (orig msg #{original_msg_idx}), analyzing separately")
                    else:
                        print(f"  Skipping: Same conversation ID (orig msg #{original_msg_idx}: '{original_subject[:35]}...')")
                        skipped_count += 1
                continue
            
            # Mark this conversation as being processed (will update result after analysis)
            processed_conversations[conversation_id] = {
                'subject': subject,
                'msg_idx': idx,
                'result': 'pending'
            }
            
            # Get all messages in the conversation
            if conversation_id not in conversation_cache:
                try:
                    conversation_messages = graph_client.get_conversation_messages(
                        conversation_id, 
                        user_email, 
                        subject_hint=subject
                    )
                    conversation_cache[conversation_id] = conversation_messages
                except Exception as e:
                    print(f"  Warning: Error fetching conversation: {e}")
                    conversation_messages = []
                    conversation_cache[conversation_id] = []
            else:
                conversation_messages = conversation_cache[conversation_id]
            
            # If conversation retrieval failed, use the sent message itself
            if not conversation_messages:
                print(f"  Using sent message itself as fallback")
                conversation_messages = [sent_msg]
            
            print(f"  Found {len(conversation_messages)} message(s) in conversation")
            
            # Get the latest message for date/metadata
            latest_message = get_latest_message_in_conversation(conversation_messages, user_email)
            if not latest_message:
                skipped_count += 1
                continue
            
            # ================================================================
            # CHECK IF EMAIL IS MARKED AS "DEALT WITH"
            # ================================================================
            latest_message_id = latest_message.get("id", "")
            if EXCLUSION_CHECKER_AVAILABLE and latest_message_id:
                if is_email_instance_excluded(conversation_id, latest_message_id, user_email):
                    print(f"Skipping: Marked as 'dealt with'")
                    skipped_count += 1
                    if conversation_id in processed_conversations:
                        processed_conversations[conversation_id]['result'] = 'dealt_with'
                    continue
            
            # Get latest message info
            latest_from = latest_message.get("from", {}).get("emailAddress", {}).get("address", "N/A")
            is_from_user = is_message_from_user(latest_message, user_email)
            latest_msg_date = get_message_datetime(latest_message)
            
            print(f"  Latest msg - From: {latest_from}, Date: {latest_msg_date.strftime('%Y-%m-%d %H:%M')}, IsFromUser: {is_from_user}")
            
            # Get the date of the latest message
            latest_date_str = latest_message.get("sentDateTime") or latest_message.get("receivedDateTime")
            if not latest_date_str:
                skipped_count += 1
                continue
            
            latest_date = date_parser.parse(latest_date_str)
            if latest_date.tzinfo is None:
                latest_date = latest_date.replace(tzinfo=tz.UTC)
            
            # Calculate age in days
            days_old = (now - latest_date).total_seconds() / (24 * 3600)
            
            # Get subject for analysis
            latest_subject = latest_message.get("subject", subject)
            latest_body = latest_message.get("body", {}).get("content", "") or latest_message.get("bodyPreview", "")
            
            # ================================================================
            # V2 IMPROVEMENT: Use full conversation analysis
            # ================================================================
            print(f"  Analyzing FULL conversation thread with AI...")
            needs_action, action_type, ai_reason, confidence = analyzer.analyze_conversation_for_action(
                conversation_messages, user_email, latest_subject
            )
            
            print(f"  AI Result: needs_action={needs_action}, action_type={action_type}, confidence={confidence}")
            print(f"  Reason: {ai_reason[:60]}..." if len(ai_reason) > 60 else f"  Reason: {ai_reason}")
            
            # ================================================================
            # FIX 4: Multi-thread awareness - check for related conversations
            # Don't assume closure, just provide context
            # ================================================================
            clean_subject_current = _clean_subject(latest_subject)
            related_conversations = []
            
            # Find related conversations with same cleaned subject
            for conv_id, conv_info in processed_conversations.items():
                if conv_id != conversation_id:
                    orig_subject = conv_info.get('subject', '')
                    clean_subject_other = _clean_subject(orig_subject)
                    if clean_subject_current == clean_subject_other:
                        related_conversations.append({
                            'conv_id': conv_id,
                            'subject': orig_subject,
                            'result': conv_info.get('result', 'unknown')
                        })
            
            # If related conversations exist, add contextual note (but don't assume closure)
            if related_conversations:
                closed_related = [r for r in related_conversations if 'closed' in r['result'].lower() or 'no_action' in r['result'].lower()]
                if closed_related and needs_action:
                    # Add note that related thread may be resolved, but still flag this one
                    related_subjects = [r['subject'][:40] for r in closed_related[:2]]  # Max 2 for brevity
                    note = f" Note: Related thread(s) '{', '.join(related_subjects)}...' appear resolved, but this thread still requires attention."
                    ai_reason = ai_reason + note
                    print(f"  ℹ️  Multi-thread: Found {len(related_conversations)} related conversation(s), {len(closed_related)} closed")
            
            # ================================================================
            # NEW: Closure detection - check if user's last message contains closure signals
            # This should run BEFORE instructions/info detection
            # ================================================================
            if is_from_user and needs_action:
                latest_body_lower = latest_body.lower()
                
                # Closure signals that indicate the conversation is closed
                closure_signals = [
                    "thank you", "thanks", "confirmed", "access granted", "done", "we will close",
                    "sent you the keys", "here are the credentials", "access provided",
                    "no further action needed", "all set", "looks good", "resolved", "fixed",
                    "completed", "finished", "closed", "resolved this", "taken care of"
                ]
                
                # Resolved in another thread indicators
                resolved_elsewhere_signals = [
                    "will reply to that other chain", "resolved in another thread", "handled in separate email",
                    "resolved in another email", "will reply in other thread", "answered in another chain",
                    "handled in other email", "resolved elsewhere", "will respond in other thread"
                ]
                
                has_closure = any(signal in latest_body_lower for signal in closure_signals)
                has_resolved_elsewhere = any(signal in latest_body_lower for signal in resolved_elsewhere_signals)
                
                if has_closure or has_resolved_elsewhere:
                    # User sent a closure message - mark as no action needed
                    needs_action = False
                    action_type = "closed"
                    if has_resolved_elsewhere:
                        ai_reason = f"User indicated this is resolved/handled in another thread - no action needed"
                    else:
                        ai_reason = f"User's last message contains closure signals - conversation appears resolved"
                    confidence = "high"
                    print(f"  ✓ Closure detected: User's message contains closure signals -> 'closed'")
            
            # ================================================================
            # FIX 3: Detect when user sent instructions/info
            # Override AI's "no_action" ONLY if there's an EXPLICIT confirmation request
            # EXCLUDE credential/info sharing emails - these don't need replies
            # ================================================================
            if is_from_user and not needs_action:
                # User sent last message but AI said no_action
                # Check if user's message contains instructions/info that would warrant waiting
                # BUT only if there's a question or request for confirmation
                latest_body_lower = latest_body.lower()
                subject_lower = latest_subject.lower()
                
                # EXCLUDE credential/info sharing emails - these are one-way communication
                is_credential_email = any(term in subject_lower or term in latest_body_lower[:300] for term in [
                    "credentials", "db credentials", "api key", "password", "access key",
                    "login credentials", "connection string", "database credentials",
                    "rag db credentials", "reminder tool db credentials"
                ])
                
                if is_credential_email:
                    # Credential emails are informational only - trust AI's no_action decision
                    print(f"  ✓ Credential/info sharing email - no reply expected (AI correctly marked as no_action)")
                else:
                    # Trust AI's decision - it has full context of the conversation
                    print(f"  ✓ Trusting AI decision -> 'no_action'")
            
            # ================================================================
            # FIX 2: Time-based closure detection
            # Check if conversation is inactive for 14+ days
            # ================================================================
            days_since_activity = days_old
            
            # If user sent last message and no activity for 14+ days, auto-close
            if is_from_user and days_since_activity >= 14 and needs_action:
                # Check if there's been any meaningful activity
                # If user sent last message 14+ days ago with no response, it might be closed
                needs_action = False
                action_type = "closed"
                auto_close_reason = f"Conversation inactive for {days_since_activity:.0f} days - auto-closed"
                ai_reason = auto_close_reason
                
                # Add to auto_closed list for visibility
                to_display_temp = extract_recipient_display(latest_message, "toRecipients")
                if not to_display_temp:
                    to_display_temp = extract_recipient_display(sent_msg, "toRecipients")
                to_str_temp = ", ".join(to_display_temp)
                
                web_link_temp = latest_message.get("webLink", "") or sent_msg.get("webLink", "")
                message_id_temp = latest_message.get("id", "") or sent_msg.get("id", "")
                outlook_link_temp = create_outlook_link(web_link_temp, message_id_temp)
                
                auto_closed_emails.append({
                    "subject": latest_subject,
                    "last_activity_date": latest_date,
                    "days_inactive": days_since_activity,
                    "reason": auto_close_reason,
                    "last_msg_from": latest_from,
                    "to_str": to_str_temp,
                    "web_link": outlook_link_temp
                })
                
                skipped_count += 1
                if conversation_id in processed_conversations:
                    processed_conversations[conversation_id]['result'] = f"auto_closed: {days_since_activity:.0f}d inactive"
                print(f"  -> Auto-closed: Inactive for {days_since_activity:.0f} days")
                continue
            
            # If no action needed, skip this conversation
            if not needs_action:
                skipped_count += 1
                # Update the processed_conversations with result
                if conversation_id in processed_conversations:
                    processed_conversations[conversation_id]['result'] = f"no_action: {ai_reason}"
                print(f"  -> No action needed: {ai_reason}")
                continue
            
            # Check for urgency (only if action is needed) - AI is PRIMARY, keywords are for display/context only
            ai_urgent, ai_urgency_reason = analyzer.analyze_urgency(latest_subject, latest_body)
            
            # AI makes the decision - it's smart enough to understand urgency from context
            is_urgent = ai_urgent  # Trust AI judgment
            
            # Keywords are only for display/context, not for decision-making
            has_keywords, priority_keywords = detect_priority_keywords(latest_subject, latest_body)
            
            # Build comprehensive urgency reason - prioritize AI explanation
            if ai_urgent and ai_urgency_reason:
                urgency_reason = ai_urgency_reason
                # Add keywords as additional context if available (for display only)
                if priority_keywords:
                    urgency_reason += f" (Keywords detected: {', '.join(priority_keywords)})"
            elif has_keywords and priority_keywords:
                # If AI didn't detect urgency but keywords found, show keywords for user awareness
                # But don't mark as urgent - trust AI's judgment
                urgency_reason = f"Keywords detected: {', '.join(priority_keywords)} (AI analysis: Not urgent)"
                is_urgent = False  # Don't override AI
            else:
                urgency_reason = ""
            
            # Get recipients display
            to_display = extract_recipient_display(latest_message, "toRecipients")
            if not to_display:
                to_display = extract_recipient_display(sent_msg, "toRecipients")
            to_str = ", ".join(to_display)
            
            # Get web link
            web_link = latest_message.get("webLink", "") or sent_msg.get("webLink", "")
            message_id = latest_message.get("id", "") or sent_msg.get("id", "")
            outlook_link = create_outlook_link(web_link, message_id)
            
            # Determine action needed based on AI analysis and basic factual check
            # Trust AI's action_type, but validate against who sent the last message
            
            if is_from_user:
                # User sent the last message - they're waiting for others to reply
                # This is a factual check - if user sent last message, they can't need to reply
                action_needed = "Waiting for reply"
                pending_from = to_str
                # Trust AI's reason to be correct (it should say who user is waiting for)
                if action_type == "user_reply_needed":
                    print(f"  ⚠️  Note: AI said 'user_reply_needed' but user sent last message - using 'Waiting for reply' (AI should have set action_type to 'waiting_for_others')")
            elif action_type == "user_reply_needed":
                # Someone else sent the last message AND AI says user needs to reply
                action_needed = "You need to reply"
                pending_from = latest_from
            elif action_type == "waiting_for_others":
                # AI says waiting for others, but someone else sent the last message
                # This is unusual - trust AI's judgment but log for review
                action_needed = "Waiting for reply"
                pending_from = to_str
                print(f"  ⚠️  Note: AI said 'waiting_for_others' but someone else sent last message - trusting AI's analysis")
            else:
                # For closed/no_action or unknown types, default based on who sent last message
                if is_from_user:
                    action_needed = "Waiting for reply"
                    pending_from = to_str
                else:
                    action_needed = "You need to reply"
                    pending_from = latest_from
            
            
            # Create email entry
            email_entry = {
                "subject": latest_subject,
                "last_msg_date": latest_msg_date,
                "last_msg_from": latest_from,
                "days_old": days_old,
                "action_needed": action_needed,
                "pending_from": pending_from,
                "to_str": to_str,
                "web_link": outlook_link,
                "priority_keywords": priority_keywords,
                "urgency_reason": urgency_reason,
                "ai_reason": ai_reason,
                "confidence": confidence,  # Store confidence for visual indicators
                "conversation_id": conversation_id,
                "latest_message_id": latest_message_id,
                "user_email": user_email
            }
            
            # CATEGORIZATION by age and urgency
            if is_urgent:
                urgent_emails.append(email_entry)
                if conversation_id in processed_conversations:
                    processed_conversations[conversation_id]['result'] = 'URGENT'
                print(f"  -> Added to URGENT list - {action_needed}")
            elif days_old < REPLY_WAIT_DAYS:
                recent_important.append(email_entry)
                if conversation_id in processed_conversations:
                    processed_conversations[conversation_id]['result'] = 'RECENT_IMPORTANT'
                print(f"  -> Added to RECENT IMPORTANT list ({days_old:.1f}d) - {action_needed}")
            else:
                hanging_emails.append(email_entry)
                if conversation_id in processed_conversations:
                    processed_conversations[conversation_id]['result'] = 'HANGING'
                print(f"  -> Added to HANGING list ({days_old:.1f}d) - {action_needed}")
        
        except Exception as e:
            print(f"  Error processing message: {e}")
            skipped_count += 1
            continue
    
    # Build stats
    stats = {
        "no_action": skipped_count,
        "total_processed": len(sent_messages)
    }
    
    return urgent_emails, recent_important, hanging_emails, auto_closed_emails, stats


def main():
    """Main function to analyze multiple team members' mailboxes and create draft emails."""
    from datetime import datetime
    
    # Set up file logging - creates a log file with timestamp
    log_filename = f"email_analysis_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    tee_output = TeeOutput(log_filename)
    sys.stdout = tee_output
    
    print(f"📝 Logging output to: {log_filename}")
    print(f"📅 Run started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Team members to analyze
    TEAM_MEMBERS = [
        "peter.koczanski@desri.com",
        "russell.petrella@desri.com"
    ]
    
    # Email where drafts will be created
    DRAFT_RECIPIENT_EMAIL = "arshdeep.kaur@desri.com"
    
    # Validate environment variables
    if not CLIENT_ID:
        print("Error: AZURE_CLIENT_ID not set in environment variables")
        sys.exit(1)
    
    if not OPENAI_API_KEY:
        print("Error: OPENAI_API_KEY not set in environment variables")
        sys.exit(1)
    
    # Check if using app-only auth (has client secret)
    is_app_only_auth = bool(CLIENT_SECRET)
    
    if not is_app_only_auth:
        print("Error: This script requires app-only authentication (CLIENT_SECRET must be set)")
        print("This is needed to access multiple mailboxes.")
        sys.exit(1)
    
    # SECURITY WARNING for app-only auth
    print("\n" + "="*70)
    print("⚠️  SECURITY WARNING: App-Only Authentication Mode")
    print("="*70)
    print("You are using CLIENT_SECRET which enables app-only authentication.")
    print("This mode can access ANY mailbox in your organization!")
    print("="*70 + "\n")
    
    # Initialize clients
    print("Initializing Microsoft Graph API client...")
    graph_client = GraphAPIClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    
    print("Initializing OpenAI client...")
    analyzer = EmailAnalyzer(OPENAI_API_KEY)
    
    print(f"\n📋 Analyzing {len(TEAM_MEMBERS)} team member(s)...")
    print(f"📧 Draft emails will be created in: {DRAFT_RECIPIENT_EMAIL}")
    print(f"{'='*80}\n")
    
    # Process each team member
    for idx, team_member_email in enumerate(TEAM_MEMBERS, 1):
        print(f"\n{'='*80}")
        print(f"Processing {idx}/{len(TEAM_MEMBERS)}: {team_member_email}")
        print(f"{'='*80}")
        
        try:
            # Analyze this user's mailbox
            urgent_emails, recent_important, hanging_emails, auto_closed_emails, stats = analyze_user_mailbox(
                team_member_email, graph_client, analyzer
            )
            
            # Calculate total items needing attention
            total_attention = len(urgent_emails) + len(recent_important) + len(hanging_emails)
            
            # Display results summary
            print(f"\n{'='*80}")
            print(f"📧 SUMMARY for {team_member_email}")
            print(f"{'='*80}")
            print(f"  🔴 Urgent:     {len(urgent_emails)}")
            print(f"  🟠 Recent Important:   {len(recent_important)}")
            print(f"  🔵 Hanging:              {len(hanging_emails)}")
            print(f"  ⚪ Auto-Closed:         {len(auto_closed_emails)}")
            print(f"  ✅ No action needed:               {stats.get('no_action', 0)}")
            print(f"{'='*80}\n")
            
            # Format team member name for subject and header
            team_member_name = format_sender_name(team_member_email)
            
            # Build enhanced digest with user name in header
            full_body = build_enhanced_digest(urgent_emails, recent_important, hanging_emails, auto_closed_emails, stats, team_member_name)
            
            # Create draft email in vikas's mailbox
            if total_attention > 0:
                # Build subject line with counts
                subject_parts = []
                if urgent_emails:
                    subject_parts.append(f"{len(urgent_emails)} urgent")
                if recent_important:
                    subject_parts.append(f"{len(recent_important)} important")
                if hanging_emails:
                    subject_parts.append(f"{len(hanging_emails)} hanging")
                
                subject = f"[Email Digest] {team_member_name} - {', '.join(subject_parts)} - {datetime.now().strftime('%b %d')}"
            else:
                subject = f"[Email Digest] {team_member_name} - All caught up! - {datetime.now().strftime('%b %d')}"
            
            print(f"Creating draft email in {DRAFT_RECIPIENT_EMAIL}'s mailbox...")
            
            try:
                draft = graph_client.create_draft_email(DRAFT_RECIPIENT_EMAIL, subject, full_body, DRAFT_RECIPIENT_EMAIL)
                draft_id = draft.get("id", "unknown")
                print(f"✓ Draft email created successfully!")
                print(f"  Draft ID: {draft_id}")
                print(f"  Subject: {subject}")
                print(f"  Check {DRAFT_RECIPIENT_EMAIL}'s Drafts folder in Outlook")
            except Exception as e:
                error_msg = str(e)
                print(f"✗ Error creating draft: {e}")
                
                if "403" in error_msg or "Forbidden" in error_msg or "Access is denied" in error_msg:
                    print("\n⚠ PERMISSION ERROR - Need Mail.ReadWrite permission")
                
                # Save to file as backup
                try:
                    safe_name = team_member_email.replace("@", "_at_").replace(".", "_")
                    output_filename = f"email_digest_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                    with open(output_filename, "w", encoding="utf-8") as f:
                        f.write(f"<html><head><title>{subject}</title></head><body>{full_body}</body></html>")
                    print(f"\n✓ Email content saved to: {output_filename}")
                except Exception as e3:
                    print(f"\n⚠ Could not save email to file: {e3}")
        
        except Exception as e:
            print(f"\n✗ Error processing {team_member_email}: {e}")
            print(f"  Continuing with next team member...")
            continue
    
    print(f"\n{'='*80}")
    print(f"✅ Analysis complete! Check {DRAFT_RECIPIENT_EMAIL}'s Drafts folder for all reports.")
    print(f"📅 Run completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📝 Full log saved to: {log_filename}")
    print(f"{'='*80}\n")
    
    # Close the log file and restore stdout
    sys.stdout = tee_output.terminal
    tee_output.close()
    print(f"Log file saved: {log_filename}")


if __name__ == "__main__":
    main()