
import sqlite3
import os

data_dir = os.environ.get("DATA_DIR", ".")
db_path = os.path.join(data_dir, "kumanogo.db")

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

settings = [
    ('smtp_host', 'smtp.lolipop.jp'),
    ('smtp_port', '465'),
    ('smtp_user', 'kumanomori@kumanomorikaken.co.jp'),
    ('smtp_pass', 'r9Fs1-k-2wf2W_6b'),
    ('smtp_from', 'kumanomori@kumanomorikaken.co.jp'),
    ('notification_email', 'nakamura@connect-web.jp')
]

for key, value in settings:
    cursor.execute("SELECT id FROM system_settings WHERE key = ?", (key,))
    row = cursor.fetchone()
    if row:
        cursor.execute("UPDATE system_settings SET value = ? WHERE key = ?", (value, key))
    else:
        cursor.execute("INSERT INTO system_settings (key, value) VALUES (?, ?)", (key, value))

conn.commit()
conn.close()
print("SMTP settings updated successfully.")
