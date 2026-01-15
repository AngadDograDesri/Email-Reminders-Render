"""
Webhook API Server for "Mark as Dealt With" Feature
Handles storing and checking excluded email instances in SQLite database.
"""

import os
import sys
import sqlite3
import json
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict
from flask import Flask, request, jsonify, make_response
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
CORS(app)  # Enable CORS for cross-origin requests

# Configuration
DB_PATH = os.getenv("EXCLUSIONS_DB_PATH", "excluded_instances.db")
API_KEY = os.getenv("WEBHOOK_API_KEY", None)  # Optional API key for authentication
PORT = int(os.getenv("WEBHOOK_PORT", "5000"))
HOST = os.getenv("WEBHOOK_HOST", "0.0.0.0")
AUTO_CLEANUP_DAYS = 14  # Auto-delete exclusions older than this many days


def init_database():
    """Initialize the SQLite database with the required schema."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
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
    
    # Migration: Add 'subject' column if it doesn't exist (for existing databases)
    try:
        cursor.execute("SELECT subject FROM excluded_instances LIMIT 1")
    except sqlite3.OperationalError:
        # Column doesn't exist, add it
        print("Migrating database: Adding 'subject' column...")
        cursor.execute("ALTER TABLE excluded_instances ADD COLUMN subject TEXT")
        print("✓ Migration complete: 'subject' column added")
    
    # Create indexes for faster lookups
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_conversation_message_user 
        ON excluded_instances(conversation_id, latest_message_id, user_email)
    """)
    
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_conversation_user 
        ON excluded_instances(conversation_id, user_email)
    """)
    
    conn.commit()
    conn.close()
    print(f"Database initialized at: {DB_PATH}")


def cleanup_old_exclusions():
    """
    Auto-cleanup: Delete exclusions older than AUTO_CLEANUP_DAYS (default 14 days).
    This ensures the database doesn't grow indefinitely and old "dealt with" 
    items are automatically reopened if they become relevant again.
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Calculate cutoff date (14 days ago)
        cutoff_date = (datetime.now(timezone.utc) - timedelta(days=AUTO_CLEANUP_DAYS)).isoformat()
        
        # Delete old records
        cursor.execute("""
            DELETE FROM excluded_instances
            WHERE excluded_at < ?
        """, (cutoff_date,))
        
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()
        
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


def check_api_key():
    """Check if API key is required and validate it."""
    if API_KEY:
        provided_key = request.headers.get("X-API-Key") or request.args.get("api_key")
        if provided_key != API_KEY:
            return jsonify({"error": "Invalid API key"}), 401
    return None


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
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Insert or update (using INSERT OR REPLACE to handle duplicates)
        cursor.execute("""
            INSERT OR REPLACE INTO excluded_instances 
            (conversation_id, latest_message_id, user_email, subject, excluded_at, reason)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (conversation_id, latest_message_id, user_email.lower(), subject, datetime.now(timezone.utc).isoformat(), reason))
        
        conn.commit()
        conn.close()
        
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
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT conversation_id, latest_message_id, subject, excluded_at, reason
            FROM excluded_instances
            WHERE user_email = ?
            ORDER BY excluded_at DESC
        """, (user_email.lower(),))
        
        results = cursor.fetchall()
        conn.close()
        
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
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute("""
            DELETE FROM excluded_instances
            WHERE conversation_id = ? 
            AND latest_message_id = ? 
            AND user_email = ?
        """, (conversation_id, latest_message_id, user_email.lower()))
        
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()
        
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
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM excluded_instances")
        count = cursor.fetchone()[0]
        conn.close()
        
        return jsonify({
            "status": "healthy",
            "database": DB_PATH,
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
    Database: {DB_PATH}
    Host: {HOST}
    Port: {PORT}
    API Key Required: {'Yes' if API_KEY else 'No'}
    Auto-Cleanup: {AUTO_CLEANUP_DAYS} days
    
    Endpoints:
    - POST/GET /api/mark-dealt-with (returns HTML for browser, JSON for API)
    - GET /api/check-excluded/<conversationId>/<latestMessageId>/<userEmail>
    - GET /api/exclusions/<userEmail>
    - POST/GET /api/undo-exclusion
    - GET /api/health
    
    Starting server...
    ================================================================
    """)
    
    app.run(host=HOST, port=PORT, debug=True)

