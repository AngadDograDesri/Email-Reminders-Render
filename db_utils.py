"""
Database utility module for handling PostgreSQL (production) and SQLite (local dev).
Automatically detects the environment based on DATABASE_URL.
"""

import os
from urllib.parse import urlparse
from typing import Optional, Tuple

# Detect database type from environment
DATABASE_URL = os.getenv("DATABASE_URL")
IS_POSTGRES = DATABASE_URL and DATABASE_URL.startswith("postgres")

if IS_POSTGRES:
    import psycopg2
    from psycopg2.extras import RealDictCursor
else:
    import sqlite3


def get_connection():
    """Get a database connection (PostgreSQL or SQLite depending on environment)."""
    if IS_POSTGRES:
        return psycopg2.connect(DATABASE_URL)
    else:
        # Local development with SQLite
        db_path = os.getenv("EXCLUSIONS_DB_PATH", "excluded_instances.db")
        return sqlite3.connect(db_path)


def init_database():
    """Initialize the database with the required schema."""
    conn = get_connection()
    cursor = conn.cursor()
    
    if IS_POSTGRES:
        # PostgreSQL schema
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS excluded_instances (
                id SERIAL PRIMARY KEY,
                conversation_id TEXT NOT NULL,
                latest_message_id TEXT NOT NULL,
                user_email TEXT NOT NULL,
                excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                reason TEXT,
                UNIQUE(conversation_id, latest_message_id, user_email)
            )
        """)
        
        # Create indexes for faster lookups
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_conversation_message_user 
            ON excluded_instances(conversation_id, latest_message_id, user_email)
        """)
        
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_conversation_user 
            ON excluded_instances(conversation_id, user_email)
        """)
    else:
        # SQLite schema
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
    
    db_type = "PostgreSQL" if IS_POSTGRES else "SQLite"
    print(f"Database initialized ({db_type})")


def execute_query(query: str, params: tuple = (), fetch_one: bool = False, fetch_all: bool = False):
    """
    Execute a database query with automatic parameter placeholder conversion.
    
    Args:
        query: SQL query (use ? for placeholders, will be converted to %s for PostgreSQL)
        params: Query parameters
        fetch_one: If True, returns one result
        fetch_all: If True, returns all results
    
    Returns:
        Result of query or None
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    # Convert SQLite placeholders (?) to PostgreSQL placeholders (%s)
    if IS_POSTGRES:
        query = query.replace("?", "%s")
    
    try:
        cursor.execute(query, params)
        
        if fetch_one:
            result = cursor.fetchone()
            conn.close()
            return result
        elif fetch_all:
            results = cursor.fetchall()
            conn.close()
            return results
        else:
            conn.commit()
            rowcount = cursor.rowcount
            conn.close()
            return rowcount
    except Exception as e:
        conn.close()
        raise e


def check_exclusion_exists(conversation_id: str, latest_message_id: str, user_email: str) -> bool:
    """Check if an exclusion exists in the database."""
    query = """
        SELECT id FROM excluded_instances
        WHERE conversation_id = ? 
        AND latest_message_id = ? 
        AND user_email = ?
    """
    result = execute_query(query, (conversation_id, latest_message_id, user_email.lower()), fetch_one=True)
    return result is not None


def add_exclusion(conversation_id: str, latest_message_id: str, user_email: str, reason: Optional[str] = None):
    """Add an exclusion to the database."""
    from datetime import datetime, timezone
    
    if IS_POSTGRES:
        # PostgreSQL uses ON CONFLICT
        query = """
            INSERT INTO excluded_instances 
            (conversation_id, latest_message_id, user_email, excluded_at, reason)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT (conversation_id, latest_message_id, user_email) 
            DO UPDATE SET excluded_at = EXCLUDED.excluded_at, reason = EXCLUDED.reason
        """
    else:
        # SQLite uses INSERT OR REPLACE
        query = """
            INSERT OR REPLACE INTO excluded_instances 
            (conversation_id, latest_message_id, user_email, excluded_at, reason)
            VALUES (?, ?, ?, ?, ?)
        """
    
    execute_query(query, (
        conversation_id,
        latest_message_id,
        user_email.lower(),
        datetime.now(timezone.utc).isoformat(),
        reason
    ))


def remove_exclusion(conversation_id: str, latest_message_id: str, user_email: str) -> int:
    """Remove an exclusion from the database. Returns number of rows deleted."""
    query = """
        DELETE FROM excluded_instances
        WHERE conversation_id = ? 
        AND latest_message_id = ? 
        AND user_email = ?
    """
    return execute_query(query, (conversation_id, latest_message_id, user_email.lower()))


def get_exclusions_for_user(user_email: str):
    """Get all exclusions for a specific user."""
    query = """
        SELECT conversation_id, latest_message_id, excluded_at, reason
        FROM excluded_instances
        WHERE user_email = ?
        ORDER BY excluded_at DESC
    """
    return execute_query(query, (user_email.lower(),), fetch_all=True)


def cleanup_old_exclusions(days: int = 14) -> int:
    """Delete exclusions older than specified days. Returns number of deleted rows."""
    from datetime import datetime, timedelta, timezone
    
    cutoff_date = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    query = """
        DELETE FROM excluded_instances
        WHERE excluded_at < ?
    """
    return execute_query(query, (cutoff_date,))


def get_total_exclusions() -> int:
    """Get total count of exclusions in database."""
    query = "SELECT COUNT(*) FROM excluded_instances"
    result = execute_query(query, fetch_one=True)
    return result[0] if result else 0

