import sqlite3
import os

db_path = 'kumanogo.db'
if not os.path.exists(db_path):
    print(f"Error: {db_path} not found.")
    exit(1)

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    cursor.execute("SELECT username, full_name, is_admin FROM users")
    users = cursor.fetchall()
    print("--- 登録ユーザー一覧 ---")
    for u in users:
        role = "管理者" if u[2] else "一般"
        print(f"ユーザー名: {u[0]}, 名前: {u[1]}, 権限: {role}")
except Exception as e:
    print(f"エラー: {e}")
finally:
    conn.close()
