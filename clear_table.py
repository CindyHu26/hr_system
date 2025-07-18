import sqlite3
from pathlib import Path

# --- Database Path ---
# This ensures the script always finds the correct database file
# when run from the same directory.
SCRIPT_DIR = Path(__file__).parent
DB_NAME = SCRIPT_DIR / "hr_system.db"

def clear_table_interactively():
    """
    Interactively prompts the user to select and clear a table
    in the specified SQLite database.
    """
    if not DB_NAME.exists():
        print(f"Error: Database file not found at '{DB_NAME}'")
        return

    try:
        # Connect to the database
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        # --- Step 1: Get and display all table names ---
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = [row[0] for row in cursor.fetchall()]
        
        # Filter out internal sqlite tables
        tables = [t for t in tables if not t.startswith('sqlite_')]

        if not tables:
            print("No tables found in the database.")
            return

        print("--- Available Tables in Database ---")
        for table in tables:
            print(f"- {table}")
        print("------------------------------------")

        # --- Step 2: Ask user for input ---
        table_to_clear = input("Please type the name of the table you want to clear: ").strip()

        if table_to_clear not in tables:
            print(f"Error: Table '{table_to_clear}' does not exist. Aborting.")
            return

        # --- Step 3: Require final confirmation ---
        print("\n" + "="*40)
        print(f"WARNING: You are about to delete ALL data from the '{table_to_clear}' table.")
        print("This operation is IRREVERSIBLE.")
        print("="*40 + "\n")
        
        confirmation = input(f"Are you absolutely sure? Type 'YES' to proceed: ")

        if confirmation != "YES":
            print("Confirmation not received. Aborting.")
            return

        # --- Step 4: Execute the delete operation ---
        print(f"\nClearing data from '{table_to_clear}'...")
        cursor.execute(f"DELETE FROM {table_to_clear};")
        
        # Get the number of deleted rows
        deleted_rows = cursor.rowcount
        
        # Commit the changes to the database
        conn.commit()
        
        print(f"Success! Deleted {deleted_rows} rows from the '{table_to_clear}' table.")

    except sqlite3.Error as e:
        print(f"A database error occurred: {e}")
    finally:
        # --- Step 5: Close the connection ---
        if conn:
            conn.close()
            print("Database connection closed.")

if __name__ == "__main__":
    clear_table_interactively()