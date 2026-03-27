
import sqlite3
conn = sqlite3.connect('kumanogo.db')
cursor = conn.cursor()
cursor.execute("SELECT * FROM system_settings;")
rows = cursor.fetchall()
for row in rows:
    print(row)
conn.close()
