"""
Database Initialization Script for Excluded Instances
Run this once to set up the database schema (PostgreSQL or SQLite).
"""

import os
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

from db_utils import init_database, get_total_exclusions, IS_POSTGRES


def setup_database():
    """Initialize the database with the required schema."""
    db_type = "PostgreSQL" if IS_POSTGRES else "SQLite"
    print(f"Initializing {db_type} database...")
    
    try:
        init_database()
        count = get_total_exclusions()
        
        print(f"✓ Database initialized successfully!")
        print(f"  - Database Type: {db_type}")
        print(f"  - Table: excluded_instances")
        print(f"  - Existing records: {count}")
        return True
    except Exception as e:
        print(f"✗ Error: Database initialization failed - {str(e)}")
        return False


if __name__ == "__main__":
    print("=" * 60)
    print("Excluded Instances Database Initialization")
    print("=" * 60)
    
    success = setup_database()
    
    if success:
        print("\n✓ Database is ready to use!")
        print(f"\nYou can now:")
        print(f"  1. Start the API server: python mark_dealt_with_api.py")
        print(f"  2. Use exclusion_checker.py in your analysis script")
    else:
        print("\n✗ Database initialization failed!")
        sys.exit(1)