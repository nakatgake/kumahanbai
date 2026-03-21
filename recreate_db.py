from database import engine, Base
import models
import os
import sqlite3

# Force delete the file if it exists
if os.path.exists('kumanogo.db'):
    os.remove('kumanogo.db')

print("Creating tables...")
Base.metadata.create_all(bind=engine)

print("Verifying schema...")
conn = sqlite3.connect('kumanogo.db')
cursor = conn.cursor()
cursor.execute("PRAGMA table_info(customers)")
cols = [row[1] for row in cursor.fetchall()]
print(f"Columns in customers: {cols}")
conn.close()
