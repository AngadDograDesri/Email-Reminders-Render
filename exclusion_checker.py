"""
Helper module for checking if email instances are excluded.
Can be imported by the main analysis script when ready to integrate.
"""

import os
import sqlite3
import requests
from typing import Optional
from dotenv import load_dotenv

# Load environment variables
try:
    load_dotenv()
except ImportError:
    pass

# Configuration
DB_PATH = os.getenv("EXCLUSIONS_DB_PATH", "excluded_instances.db")
WEBHOOK_API_URL = os.getenv("WEBHOOK_API_URL", "http://localhost:5000")
USE_API = os.getenv("USE_EXCLUSION_API", "false").lower() == "true"  # Set to "true" to use API instead of direct DB


def is_email_instance_excluded(conversation_id: str, latest_message_id: str, user_email: str) -> bool:
    """
    Check if a specific email instance is marked as "dealt with".
    
    Args:
        conversation_id: The conversation ID from Microsoft Graph API
        latest_message_id: The ID of the latest message in the conversation
        user_email: The email address of the user
    
    Returns:
        True if the instance is excluded, False otherwise
    """
    if not conversation_id or not latest_message_id or not user_email:
        return False
    
    if USE_API:
        return _check_via_api(conversation_id, latest_message_id, user_email)
    else:
        return _check_via_db(conversation_id, latest_message_id, user_email)


def _check_via_api(conversation_id: str, latest_message_id: str, user_email: str) -> bool:
    """Check exclusion via webhook API."""
    try:
        url = f"{WEBHOOK_API_URL}/api/check-excluded/{conversation_id}/{latest_message_id}/{user_email}"
        
        # Add API key if configured
        headers = {}
        api_key = os.getenv("WEBHOOK_API_KEY")
        if api_key:
            headers["X-API-Key"] = api_key
        
        response = requests.get(url, headers=headers, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            return data.get("excluded", False)
        else:
            # If API is unavailable, fall back to direct DB access
            print(f"  Warning: API check failed ({response.status_code}), falling back to direct DB")
            return _check_via_db(conversation_id, latest_message_id, user_email)
            
    except Exception as e:
        # If API call fails, fall back to direct DB access
        print(f"  Warning: API check error ({str(e)}), falling back to direct DB")
        return _check_via_db(conversation_id, latest_message_id, user_email)


def _check_via_db(conversation_id: str, latest_message_id: str, user_email: str) -> bool:
    """Check exclusion via direct database access."""
    if not os.path.exists(DB_PATH):
        return False
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT id FROM excluded_instances
            WHERE conversation_id = ? 
            AND latest_message_id = ? 
            AND user_email = ?
        """, (conversation_id, latest_message_id, user_email.lower()))
        
        result = cursor.fetchone()
        conn.close()
        
        return result is not None
        
    except Exception as e:
        print(f"  Warning: Database check error: {str(e)}")
        return False


def mark_as_dealt_with(conversation_id: str, latest_message_id: str, user_email: str, reason: Optional[str] = None) -> bool:
    """
    Mark an email instance as dealt with (programmatically).
    
    Args:
        conversation_id: The conversation ID
        latest_message_id: The ID of the latest message
        user_email: The user's email address
        reason: Optional reason for exclusion
    
    Returns:
        True if successful, False otherwise
    """
    if USE_API:
        return _mark_via_api(conversation_id, latest_message_id, user_email, reason)
    else:
        return _mark_via_db(conversation_id, latest_message_id, user_email, reason)


def _mark_via_api(conversation_id: str, latest_message_id: str, user_email: str, reason: Optional[str] = None) -> bool:
    """Mark exclusion via webhook API."""
    try:
        url = f"{WEBHOOK_API_URL}/api/mark-dealt-with"
        
        headers = {"Content-Type": "application/json"}
        api_key = os.getenv("WEBHOOK_API_KEY")
        if api_key:
            headers["X-API-Key"] = api_key
        
        data = {
            "conversationId": conversation_id,
            "latestMessageId": latest_message_id,
            "userEmail": user_email,
        }
        if reason:
            data["reason"] = reason
        
        response = requests.post(url, json=data, headers=headers, timeout=5)
        return response.status_code == 200
        
    except Exception as e:
        print(f"  Warning: API mark error: {str(e)}")
        return False


def _mark_via_db(conversation_id: str, latest_message_id: str, user_email: str, reason: Optional[str] = None) -> bool:
    """Mark exclusion via direct database access."""
    try:
        from datetime import datetime
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Ensure table exists
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS excluded_instances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conversation_id TEXT NOT NULL,
                latest_message_id TEXT NOT NULL,
                user_email TEXT NOT NULL,
                excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                reason TEXT,
                UNIQUE(conversation_id, latest_message_id, user_email)
            )
        """)
        
        cursor.execute("""
            INSERT OR REPLACE INTO excluded_instances 
            (conversation_id, latest_message_id, user_email, excluded_at, reason)
            VALUES (?, ?, ?, ?, ?)
        """, (conversation_id, latest_message_id, user_email.lower(), datetime.utcnow().isoformat(), reason))
        
        conn.commit()
        conn.close()
        return True
        
    except Exception as e:
        print(f"  Warning: Database mark error: {str(e)}")
        return False

