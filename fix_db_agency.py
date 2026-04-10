import sqlite3

db_path = '/var/www/kumanomori/kumanogo.db'

try:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("ALTER TABLE agency_orders ADD COLUMN invoice_id INTEGER;")
    conn.commit()
    print("SUCCESS: Kolumn added")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("INFO: Column already exists, all good.")
    else:
        print(f"ERROR: {e}")
finally:
    if 'conn' in locals():
        conn.close()
