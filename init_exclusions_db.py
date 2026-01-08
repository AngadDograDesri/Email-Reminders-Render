"""
Database Initialization Script for Excluded Instances
Run this once to set up the SQLite database schema.
"""

import os
import sqlite3
import sys

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("Warning: python-dotenv not installed. Make sure to set environment variables manually.")

DB_PATH = os.getenv("EXCLUSIONS_DB_PATH", "excluded_instances.db")


def init_database():
    """Initialize the SQLite database with the required schema."""
    print(f"Initializing database at: {DB_PATH}")
    
    # Create directory if it doesn't exist
    db_dir = os.path.dirname(os.path.abspath(DB_PATH))
    if db_dir and not os.path.exists(db_dir):
        os.makedirs(db_dir, exist_ok=True)
        print(f"Created directory: {db_dir}")
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Create main table
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
    
    # Verify table was created
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='excluded_instances'")
    table_exists = cursor.fetchone() is not None
    
    if table_exists:
        cursor.execute("SELECT COUNT(*) FROM excluded_instances")
        count = cursor.fetchone()[0]
        print(f"✓ Database initialized successfully!")
        print(f"  - Table: excluded_instances")
        print(f"  - Existing records: {count}")
    else:
        print("✗ Error: Table was not created")
    
    conn.close()
    return table_exists


if __name__ == "__main__":
    print("=" * 60)
    print("Excluded Instances Database Initialization")
    print("=" * 60)
    
    success = init_database()
    
    if success:
        print("\n✓ Database is ready to use!")
        print(f"\nYou can now:")
        print(f"  1. Start the API server: python mark_dealt_with_api.py")
        print(f"  2. Use exclusion_checker.py in your analysis script")
    else:
        print("\n✗ Database initialization failed!")
        sys.exit(1)