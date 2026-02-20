"""
Webhook API Server for "Mark as Dealt With" Feature
Handles storing and checking excluded email instances.
Supports SQLite (default) or PostgreSQL when DATABASE_URL is set.
OAuth (Azure) and encrypted token storage for delegated mailbox access.
"""

import os
import sys
import sqlite3
import json
import base64
import hashlib
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, Any, List, Tuple
from contextlib import contextmanager
from flask import Flask, request, jsonify, make_response, redirect, session
from flask_cors import CORS

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("Warning: python-dotenv not installed. Make sure to set environment variables manually.")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(24).hex())
CORS(app, supports_credentials=True)

# Configuration
DATABASE_URL = os.getenv("DATABASE_URL")
DB_PATH = os.getenv("EXCLUSIONS_DB_PATH", "excluded_instances.db")
API_KEY = os.getenv("WEBHOOK_API_KEY", None)
PORT = int(os.getenv("WEBHOOK_PORT", "5000"))
HOST = os.getenv("WEBHOOK_HOST", "0.0.0.0")
AUTO_CLEANUP_DAYS = 14

# OAuth / token storage
TOKEN_ENCRYPTION_KEY = os.getenv("TOKEN_ENCRYPTION_KEY")
INTERNAL_API_KEY = os.getenv("INTERNAL_API_KEY")
# Set to "true" or "1" to store refresh tokens without encryption (avoids decrypt_failed; use only if DB is not exposed).
STORE_TOKENS_PLAINTEXT = os.getenv("STORE_TOKENS_PLAINTEXT", "").strip().lower() in ("1", "true", "yes")
REDIRECT_URI = os.getenv("REDIRECT_URI")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")

USE_POSTGRES = bool(DATABASE_URL)


def get_connection():
    """Return a database connection (SQLite or Postgres)."""
    if USE_POSTGRES:
        import psycopg2
        return psycopg2.connect(DATABASE_URL)
    return sqlite3.connect(DB_PATH)


@contextmanager
def db_cursor(commit: bool = True):
    """Context manager for DB cursor. Use ? for SQLite, %s for Postgres."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        yield cur
        if commit:
            conn.commit()
    finally:
        cur.close()
        conn.close()


def _param_style() -> str:
    """Return placeholder style: ? for SQLite, %s for Postgres."""
    return "?" if not USE_POSTGRES else "%s"


def init_database():
    """Initialize database (SQLite or Postgres) with required schema."""
    with db_cursor(commit=True) as cur:
        if USE_POSTGRES:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS excluded_instances (
                    id SERIAL PRIMARY KEY,
                    conversation_id TEXT NOT NULL,
                    latest_message_id TEXT NOT NULL,
                    user_email TEXT NOT NULL,
                    subject TEXT,
                    excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    reason TEXT,
                    UNIQUE(conversation_id, latest_message_id, user_email)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS user_tokens (
                    user_email TEXT PRIMARY KEY,
                    encrypted_refresh_token BYTEA NOT NULL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
        else:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS excluded_instances (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conversation_id TEXT NOT NULL,
                    latest_message_id TEXT NOT NULL,
                    user_email TEXT NOT NULL,
                    subject TEXT,
                    excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    reason TEXT,
                    UNIQUE(conversation_id, latest_message_id, user_email)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS user_tokens (
                    user_email TEXT PRIMARY KEY,
                    encrypted_refresh_token BLOB NOT NULL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

        # Migration: Add 'subject' to excluded_instances if missing (SQLite only)
        if not USE_POSTGRES:
            try:
                cur.execute("SELECT subject FROM excluded_instances LIMIT 1")
            except sqlite3.OperationalError:
                cur.execute("ALTER TABLE excluded_instances ADD COLUMN subject TEXT")

        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_conversation_message_user ON excluded_instances(conversation_id, latest_message_id, user_email)"
        )
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_conversation_user ON excluded_instances(conversation_id, user_email)"
        )
    print(f"Database initialized: {'Postgres' if USE_POSTGRES else DB_PATH}")


def cleanup_old_exclusions():
    """
    Auto-cleanup: Delete exclusions older than AUTO_CLEANUP_DAYS (default 14 days).
    """
    try:
        p = _param_style()
        cutoff_date = (datetime.now(timezone.utc) - timedelta(days=AUTO_CLEANUP_DAYS)).isoformat()
        with db_cursor() as cur:
            cur.execute(
                f"DELETE FROM excluded_instances WHERE excluded_at < {p}",
                (cutoff_date,)
            )
            deleted_count = cur.rowcount
        if deleted_count > 0:
            print(f"✓ Auto-cleanup: Removed {deleted_count} exclusion(s) older than {AUTO_CLEANUP_DAYS} days")
        else:
            print(f"✓ Auto-cleanup: No old exclusions to remove")
        return deleted_count
    except Exception as e:
        print(f"⚠ Auto-cleanup error: {str(e)}")
        return 0


def generate_success_html(user_email: str, subject: str = "") -> str:
    """Generate a nice HTML success page instead of raw JSON."""
    import html
    escaped_subject = html.escape(subject[:80]) + ("..." if len(subject) > 80 else "") if subject else ""
    subject_html = f'<p style="color:#333; font-size:14px; margin-top:15px;"><strong>Subject:</strong> {escaped_subject}</p>' if subject else ""
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Marked as Dealt With</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: 0;
                padding: 20px;
                box-sizing: border-box;
            }}
            .container {{
                background: white;
                border-radius: 16px;
                padding: 40px 50px;
                text-align: center;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                max-width: 450px;
                width: 100%;
            }}
            .checkmark {{
                width: 80px;
                height: 80px;
                background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
                border-radius: 50%;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: 0 auto 25px;
                box-shadow: 0 8px 25px rgba(76, 175, 80, 0.4);
            }}
            .checkmark svg {{
                width: 45px;
                height: 45px;
                fill: white;
            }}
            h1 {{
                color: #2e7d32;
                margin: 0 0 15px 0;
                font-size: 26px;
                font-weight: 600;
            }}
            p {{
                color: #666;
                margin: 0 0 10px 0;
                font-size: 15px;
                line-height: 1.5;
            }}
            .email {{
                color: #1a237e;
                font-weight: 600;
                background: #e8eaf6;
                padding: 3px 10px;
                border-radius: 4px;
                display: inline-block;
                margin-top: 5px;
            }}
            .close-hint {{
                margin-top: 25px;
                padding-top: 20px;
                border-top: 1px solid #eee;
                color: #999;
                font-size: 13px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="checkmark">
                <svg viewBox="0 0 24 24">
                    <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                </svg>
            </div>
            <h1>Marked as Dealt With!</h1>
            <p>This email has been successfully marked as dealt with.</p>
            <p>It will be <strong>skipped</strong> in future digests until new messages arrive.</p>
            {subject_html}
            <p class="email">{user_email}</p>
            <p class="close-hint">You can close this tab now.</p>
        </div>
    </body>
    </html>
    """


def generate_error_html(error_message: str) -> str:
    """Generate a nice HTML error page."""
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Error</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                background: linear-gradient(135deg, #ff6b6b 0%, #c0392b 100%);
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: 0;
                padding: 20px;
                box-sizing: border-box;
            }}
            .container {{
                background: white;
                border-radius: 16px;
                padding: 40px 50px;
                text-align: center;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                max-width: 450px;
                width: 100%;
            }}
            .error-icon {{
                width: 80px;
                height: 80px;
                background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
                border-radius: 50%;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: 0 auto 25px;
                box-shadow: 0 8px 25px rgba(244, 67, 54, 0.4);
            }}
            .error-icon svg {{
                width: 45px;
                height: 45px;
                fill: white;
            }}
            h1 {{
                color: #c62828;
                margin: 0 0 15px 0;
                font-size: 26px;
                font-weight: 600;
            }}
            p {{
                color: #666;
                margin: 0;
                font-size: 15px;
                line-height: 1.5;
            }}
            .error-detail {{
                background: #ffebee;
                color: #c62828;
                padding: 12px 16px;
                border-radius: 8px;
                margin-top: 20px;
                font-size: 13px;
                word-break: break-word;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="error-icon">
                <svg viewBox="0 0 24 24">
                    <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
                </svg>
            </div>
            <h1>Something Went Wrong</h1>
            <p>Could not mark this email as dealt with.</p>
            <div class="error-detail">{error_message}</div>
        </div>
    </body>
    </html>
    """


# ----- Token encryption and OAuth -----
def _get_fernet():
    """
    Return Fernet instance for token encryption.
    Uses TOKEN_ENCRYPTION_KEY if set; otherwise derives key from INTERNAL_API_KEY
    so the same key is used everywhere (avoids decrypt_failed across instances).
    """
    from cryptography.fernet import Fernet
    raw = (TOKEN_ENCRYPTION_KEY or "").strip() if isinstance(TOKEN_ENCRYPTION_KEY, str) else TOKEN_ENCRYPTION_KEY
    if raw:
        try:
            key = raw.encode("utf-8") if isinstance(raw, str) else raw
            return Fernet(key)
        except Exception:
            pass
    # Fallback: derive from INTERNAL_API_KEY so encrypt/decrypt always use same key
    if INTERNAL_API_KEY:
        try:
            key_bytes = hashlib.sha256(INTERNAL_API_KEY.encode()).digest()
            key = base64.urlsafe_b64encode(key_bytes).decode()
            return Fernet(key.encode())
        except Exception:
            pass
    return None


def _encrypt_refresh_token(token: str) -> Optional[bytes]:
    """Encrypt refresh token for storage (or store plaintext if STORE_TOKENS_PLAINTEXT)."""
    if STORE_TOKENS_PLAINTEXT:
        return token.encode("utf-8")
    f = _get_fernet()
    if not f:
        return None
    return f.encrypt(token.encode("utf-8"))


def _decrypt_refresh_token(encrypted: bytes) -> Optional[str]:
    """
    Read refresh token from stored bytes. Always try plaintext (UTF-8) first so that
    tokens stored with STORE_TOKENS_PLAINTEXT work regardless of env at read time.
    If that fails or doesn't look like a token, try Fernet decrypt.
    """
    try:
        plain = encrypted.decode("utf-8")
        if len(plain) >= 50:
            return plain
    except Exception:
        pass
    f = _get_fernet()
    if not f:
        return None
    try:
        return f.decrypt(encrypted).decode("utf-8")
    except Exception:
        return None


def _get_stored_refresh_token(user_email: str) -> Optional[bytes]:
    """Return encrypted refresh token for user, or None."""
    p = _param_style()
    with db_cursor() as cur:
        cur.execute(f"SELECT encrypted_refresh_token FROM user_tokens WHERE user_email = {p}", (user_email.lower(),))
        row = cur.fetchone()
    return row[0] if row else None


def _store_refresh_token(user_email: str, encrypted: bytes) -> None:
    """Upsert encrypted refresh token for user."""
    p = _param_style()
    now = datetime.now(timezone.utc).isoformat()
    if USE_POSTGRES:
        with db_cursor() as cur:
            cur.execute("""
                INSERT INTO user_tokens (user_email, encrypted_refresh_token, updated_at)
                VALUES (%s, %s, %s)
                ON CONFLICT (user_email) DO UPDATE SET encrypted_refresh_token = EXCLUDED.encrypted_refresh_token, updated_at = EXCLUDED.updated_at
            """, (user_email.lower(), encrypted, now))
    else:
        with db_cursor() as cur:
            cur.execute(
                "INSERT OR REPLACE INTO user_tokens (user_email, encrypted_refresh_token, updated_at) VALUES (?, ?, ?)",
                (user_email.lower(), encrypted, now)
            )


def get_access_token_for_user(user_email: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Get a valid access token for the user by refreshing from stored refresh token.
    Returns (access_token, None) on success, or (None, hint) on failure.
    hint: "no_stored_token", "decrypt_failed", "clear_tokens_and_sign_in_again", "missing_config", "refresh_failed".
    """
    encrypted = _get_stored_refresh_token(user_email)
    if not encrypted:
        return (None, "no_stored_token")
    refresh_token = _decrypt_refresh_token(encrypted)
    if not refresh_token:
        if STORE_TOKENS_PLAINTEXT:
            return (None, "clear_tokens_and_sign_in_again")
        return (None, "decrypt_failed")
    if not AZURE_CLIENT_ID or not AZURE_CLIENT_SECRET or not AZURE_TENANT_ID:
        return (None, "missing_config")
    try:
        from msal import ConfidentialClientApplication
        authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        app = ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            authority=authority,
            client_credential=AZURE_CLIENT_SECRET,
        )
        SCOPES = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadWrite"]
        result = app.acquire_token_by_refresh_token(
            refresh_token,
            scopes=SCOPES,
        )
        if result and "access_token" in result:
            if "refresh_token" in result:
                enc_new = _encrypt_refresh_token(result["refresh_token"])
                if enc_new:
                    _store_refresh_token(user_email, enc_new)
            return (result["access_token"], None)
        return (None, "refresh_failed")
    except Exception:
        return (None, "refresh_failed")


def check_api_key():
    """Check if API key is required and validate it."""
    if API_KEY:
        provided_key = request.headers.get("X-API-Key") or request.args.get("api_key")
        if provided_key != API_KEY:
            return jsonify({"error": "Invalid API key"}), 401
    return None


def require_internal_api_key():
    """Require X-Internal-Api-Key header; return 401 response if missing or wrong."""
    if not INTERNAL_API_KEY:
        return jsonify({"error": "Internal API not configured"}), 503
    key = request.headers.get("X-Internal-Api-Key")
    if key != INTERNAL_API_KEY:
        return jsonify({"error": "Unauthorized"}), 401
    return None


# ----- OAuth and internal token endpoints -----
@app.route("/auth/start", methods=["GET"])
def auth_start():
    """Redirect user to Azure sign-in. After callback, refresh token is stored for that user."""
    if not all([AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, REDIRECT_URI]):
        return make_response(generate_error_html("OAuth not configured (missing AZURE_* or REDIRECT_URI)"), 500)
    try:
        from msal import ConfidentialClientApplication
        import secrets
        authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        app = ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            authority=authority,
            client_credential=AZURE_CLIENT_SECRET,
        )
        state = secrets.token_urlsafe(32)
        session["oauth_state"] = state
        auth_url = app.get_authorization_request_url(
            scopes=["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadWrite"],
            state=state,
            redirect_uri=REDIRECT_URI,
        )
        return redirect(auth_url)
    except Exception as e:
        return make_response(generate_error_html(str(e)), 500)


@app.route("/auth/callback", methods=["GET"])
def auth_callback():
    """Exchange code for tokens, store encrypted refresh token, show success."""
    if not all([AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, REDIRECT_URI]):
        return make_response(generate_error_html("OAuth not configured"), 500)
    state = request.args.get("state")
    if state != session.get("oauth_state"):
        return make_response(generate_error_html("Invalid state"), 400)
    code = request.args.get("code")
    if not code:
        err = request.args.get("error_description") or request.args.get("error") or "No code"
        return make_response(generate_error_html(err), 400)
    try:
        from msal import ConfidentialClientApplication
        authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        app = ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            authority=authority,
            client_credential=AZURE_CLIENT_SECRET,
        )
        result = app.acquire_token_by_authorization_code(
            code,
            scopes=["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadWrite"],
            redirect_uri=REDIRECT_URI,
        )
        if not result or "access_token" not in result:
            err_msg = "Failed to get tokens"
            if result and isinstance(result, dict):
                err = result.get("error_description") or result.get("error")
                if err:
                    err_msg = str(err)
            return make_response(generate_error_html(err_msg), 500)
        refresh_token = result.get("refresh_token")
        if not refresh_token:
            return make_response(generate_error_html("No refresh token in response"), 500)
        # Get user email from id_token or /me
        user_email = None
        if result.get("id_token_claims"):
            user_email = (result["id_token_claims"].get("preferred_username") or result["id_token_claims"].get("upn"))
        if not user_email and result.get("access_token"):
            import requests
            r = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {result['access_token']}"},
                params={"$select": "mail,userPrincipalName"},
            )
            if r.ok:
                j = r.json()
                user_email = j.get("mail") or j.get("userPrincipalName")
        if not user_email:
            return make_response(generate_error_html("Could not determine user email"), 500)
        encrypted = _encrypt_refresh_token(refresh_token)
        if not encrypted:
            return make_response(generate_error_html("Token encryption not configured (TOKEN_ENCRYPTION_KEY)"), 500)
        # Ensure we can decrypt what we encrypted (catches wrong/inconsistent key)
        if _decrypt_refresh_token(encrypted) != refresh_token:
            return make_response(
                generate_error_html("Token storage failed: encryption key problem. Please ask an admin to set TOKEN_ENCRYPTION_KEY to a valid Fernet key (44 chars) and try again."),
                500,
            )
        _store_refresh_token(user_email, encrypted)
        session.pop("oauth_state", None)
        return make_response("""
        <!DOCTYPE html><html><head><title>Success</title></head><body>
        <h1>Signed in successfully</h1>
        <p>Your mailbox is now connected for the email digest. You can close this tab.</p>
        <p>User: """ + user_email + """</p>
        </body></html>""", 200)
    except Exception as e:
        return make_response(generate_error_html(str(e)), 500)


@app.route("/api/internal/token/<path:user_email>", methods=["GET"])
def internal_token(user_email: str):
    """Return access token for user (for cron). Requires X-Internal-Api-Key header."""
    err = require_internal_api_key()
    if err:
        return err
    access_token, hint = get_access_token_for_user(user_email)
    if not access_token:
        return jsonify({"error": "No token or refresh failed", "hint": hint or "refresh_failed"}), 404
    return jsonify({"access_token": access_token})


@app.route("/api/internal/users", methods=["GET"])
def internal_users():
    """Return list of user emails that have stored tokens. Requires X-Internal-Api-Key header."""
    err = require_internal_api_key()
    if err:
        return err
    with db_cursor() as cur:
        cur.execute("SELECT user_email FROM user_tokens ORDER BY user_email")
        rows = cur.fetchall()
    users = [r[0] for r in rows]
    return jsonify({"users": users})


@app.route("/api/internal/clear-tokens", methods=["POST"])
def internal_clear_tokens():
    """
    Delete all rows in user_tokens. Use when decrypt_failed persists (e.g. key was changed).
    After calling, have every user sign in again at /auth/start.
    Requires X-Internal-Api-Key header.
    """
    err = require_internal_api_key()
    if err:
        return err
    with db_cursor() as cur:
        cur.execute("DELETE FROM user_tokens")
        deleted = cur.rowcount
    return jsonify({
        "ok": True,
        "message": "All stored tokens cleared. Users must sign in again at /auth/start.",
        "deleted_count": deleted
    })


@app.route("/api/mark-dealt-with", methods=["POST", "GET"])
def mark_dealt_with():
    """
    Mark a specific email instance as dealt with.
    
    Accepts:
    - conversationId (required)
    - latestMessageId (required)
    - userEmail (required)
    - subject (optional) - email subject for convenience
    - reason (optional)
    
    Can be called via GET (query params) or POST (JSON body)
    Returns: Nice HTML page (for browser) or JSON (for API calls)
    """
    # Check API key if configured
    auth_error = check_api_key()
    if auth_error:
        return auth_error
    
    # Determine if this is a browser request (wants HTML) or API request (wants JSON)
    wants_html = 'text/html' in request.headers.get('Accept', '')
    # For GET requests from browser clicks, default to HTML
    if request.method == "GET" and not request.headers.get('X-Requested-With'):
        wants_html = True
    
    # Get parameters from either GET query params or POST JSON body
    if request.method == "POST":
        data = request.get_json() or {}
        conversation_id = data.get("conversationId") or data.get("conversation_id")
        latest_message_id = data.get("latestMessageId") or data.get("latest_message_id")
        user_email = data.get("userEmail") or data.get("user_email")
        subject = data.get("subject", "")
        reason = data.get("reason", "")
    else:  # GET
        conversation_id = request.args.get("conversationId") or request.args.get("conversation_id")
        latest_message_id = request.args.get("latestMessageId") or request.args.get("latest_message_id")
        user_email = request.args.get("userEmail") or request.args.get("user_email")
        subject = request.args.get("subject", "")
        reason = request.args.get("reason", "")
    
    # Validate required parameters
    if not conversation_id or not latest_message_id or not user_email:
        error_msg = "Missing required parameters: conversationId, latestMessageId, and userEmail are required"
        if wants_html:
            response = make_response(generate_error_html(error_msg), 400)
            response.headers['Content-Type'] = 'text/html'
            return response
        return jsonify({
            "success": False,
            "error": error_msg
        }), 400
    
    p = _param_style()
    try:
        if USE_POSTGRES:
            with db_cursor() as cur:
                cur.execute("""
                    INSERT INTO excluded_instances 
                    (conversation_id, latest_message_id, user_email, subject, excluded_at, reason)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (conversation_id, latest_message_id, user_email)
                    DO UPDATE SET subject = EXCLUDED.subject, excluded_at = EXCLUDED.excluded_at, reason = EXCLUDED.reason
                """, (conversation_id, latest_message_id, user_email.lower(), subject, datetime.now(timezone.utc).isoformat(), reason))
        else:
            with db_cursor() as cur:
                cur.execute("""
                    INSERT OR REPLACE INTO excluded_instances 
                    (conversation_id, latest_message_id, user_email, subject, excluded_at, reason)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (conversation_id, latest_message_id, user_email.lower(), subject, datetime.now(timezone.utc).isoformat(), reason))
        
        # Return nice HTML page for browser clicks
        if wants_html:
            response = make_response(generate_success_html(user_email, subject), 200)
            response.headers['Content-Type'] = 'text/html'
            return response
        
        # Return JSON for API calls
        return jsonify({
            "success": True,
            "message": "Email instance marked as dealt with",
            "data": {
                "conversationId": conversation_id,
                "latestMessageId": latest_message_id,
                "userEmail": user_email
            }
        }), 200
        
    except Exception as e:
        error_msg = f"Database error: {str(e)}"
        if wants_html:
            response = make_response(generate_error_html(error_msg), 500)
            response.headers['Content-Type'] = 'text/html'
            return response
        return jsonify({
            "success": False,
            "error": error_msg
        }), 500


@app.route("/api/check-excluded/<conversation_id>/<latest_message_id>/<user_email>", methods=["GET"])
def check_excluded(conversation_id: str, latest_message_id: str, user_email: str):
    """
    Check if a specific email instance is excluded.
    
    Returns: {"excluded": true/false}
    """
    # Check API key if configured
    auth_error = check_api_key()
    if auth_error:
        return auth_error
    
    p = _param_style()
    try:
        with db_cursor() as cur:
            cur.execute(
                f"SELECT id FROM excluded_instances WHERE conversation_id = {p} AND latest_message_id = {p} AND user_email = {p}",
                (conversation_id, latest_message_id, user_email.lower())
            )
            result = cur.fetchone()
        excluded = result is not None
        
        return jsonify({
            "excluded": excluded,
            "conversationId": conversation_id,
            "latestMessageId": latest_message_id,
            "userEmail": user_email
        }), 200
        
    except Exception as e:
        return jsonify({
            "error": f"Database error: {str(e)}"
        }), 500


@app.route("/api/exclusions/<user_email>", methods=["GET"])
def list_exclusions(user_email: str):
    """
    List all exclusions for a specific user.
    Optional: for admin/debugging purposes.
    """
    # Check API key if configured
    auth_error = check_api_key()
    if auth_error:
        return auth_error
    
    p = _param_style()
    try:
        with db_cursor() as cur:
            cur.execute(
                f"SELECT conversation_id, latest_message_id, subject, excluded_at, reason FROM excluded_instances WHERE user_email = {p} ORDER BY excluded_at DESC",
                (user_email.lower(),)
            )
            results = cur.fetchall()
        exclusions = []
        for row in results:
            exclusions.append({
                "conversationId": row[0],
                "latestMessageId": row[1],
                "subject": row[2],
                "excludedAt": row[3],
                "reason": row[4]
            })
        
        return jsonify({
            "userEmail": user_email,
            "count": len(exclusions),
            "exclusions": exclusions
        }), 200
        
    except Exception as e:
        return jsonify({
            "error": f"Database error: {str(e)}"
        }), 500


@app.route("/api/undo-exclusion", methods=["POST", "GET"])
def undo_exclusion():
    """
    Remove an exclusion (undo "mark as dealt with").
    
    Accepts: conversationId, latestMessageId, userEmail
    """
    # Check API key if configured
    auth_error = check_api_key()
    if auth_error:
        return auth_error
    
    # Get parameters
    if request.method == "POST":
        data = request.get_json() or {}
        conversation_id = data.get("conversationId") or data.get("conversation_id")
        latest_message_id = data.get("latestMessageId") or data.get("latest_message_id")
        user_email = data.get("userEmail") or data.get("user_email")
    else:  # GET
        conversation_id = request.args.get("conversationId") or request.args.get("conversation_id")
        latest_message_id = request.args.get("latestMessageId") or request.args.get("latest_message_id")
        user_email = request.args.get("userEmail") or request.args.get("user_email")
    
    if not conversation_id or not latest_message_id or not user_email:
        return jsonify({
            "success": False,
            "error": "Missing required parameters: conversationId, latestMessageId, and userEmail are required"
        }), 400
    
    p = _param_style()
    try:
        with db_cursor() as cur:
            cur.execute(
                f"DELETE FROM excluded_instances WHERE conversation_id = {p} AND latest_message_id = {p} AND user_email = {p}",
                (conversation_id, latest_message_id, user_email.lower())
            )
            deleted_count = cur.rowcount
        if deleted_count > 0:
            return jsonify({
                "success": True,
                "message": "Exclusion removed"
            }), 200
        else:
            return jsonify({
                "success": False,
                "error": "Exclusion not found"
            }), 404
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": f"Database error: {str(e)}"
        }), 500


@app.route("/api/health", methods=["GET"])
def health_check():
    """Health check endpoint."""
    try:
        with db_cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM excluded_instances")
            count = cur.fetchone()[0]
        return jsonify({
            "status": "healthy",
            "database": "postgres" if USE_POSTGRES else DB_PATH,
            "total_exclusions": count
        }), 200
    except Exception as e:
        return jsonify({
            "status": "unhealthy",
            "error": str(e)
        }), 500


if __name__ == "__main__":
    # Initialize database on startup
    init_database()
    
    # Auto-cleanup old exclusions (14+ days old)
    cleanup_old_exclusions()
    
    print(f"""
    ================================================================
    Mark as Dealt With API Server
    ================================================================
    Database: {'Postgres' if USE_POSTGRES else DB_PATH}
    Host: {HOST}
    Port: {PORT}
    API Key Required: {'Yes' if API_KEY else 'No'}
    Auto-Cleanup: {AUTO_CLEANUP_DAYS} days
    OAuth: {'Yes' if all([AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, REDIRECT_URI]) else 'No'}
    
    Endpoints:
    - GET /auth/start (OAuth sign-in)
    - GET /auth/callback (OAuth callback)
    - GET /api/internal/token/<user_email> (internal; X-Internal-Api-Key)
    - GET /api/internal/users (internal; X-Internal-Api-Key)
    - POST/GET /api/mark-dealt-with (returns HTML for browser, JSON for API)
    - GET /api/check-excluded/<conversationId>/<latestMessageId>/<userEmail>
    - GET /api/exclusions/<userEmail>
    - POST/GET /api/undo-exclusion
    - GET /api/health
    
    Starting server...
    ================================================================
    """)
    
    app.run(host=HOST, port=PORT, debug=True)

