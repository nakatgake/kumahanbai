import sqlite3
import os

db_path = 'kumanogo.db'
backup_path = 'kumanogo.db.before_wipe'

if not os.path.exists(db_path):
    print(f"Error: {db_path} not found.")
    exit(1)

# 1. バックアップの作成
import shutil
shutil.copy2(db_path, backup_path)
print(f"Safety backup created: {backup_path}")

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    print("--- データベース全データ初期化開始 ---")
    
    # 削除対象のテーブルリスト (users 以外)
    tables_to_wipe = [
        'customers', 'products', 'locations', 'product_stocks',
        'quotations', 'quotation_items', 'orders', 'invoices',
        'agency_orders', 'agency_order_items', 'notifications',
        'system_settings'
    ]
    
    for table in tables_to_wipe:
        try:
            cursor.execute(f"DELETE FROM {table}")
            # SQLite の ID カウントをリセット
            cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{table}'")
            print(f"  - {table} を空にしました。")
        except sqlite3.OperationalError as e:
            print(f"  - {table} は存在しません（スキップ）。")

    # 初期拠点の再登録 (システム動作に必須なため)
    cursor.execute("INSERT INTO locations (name, is_default) VALUES (?, ?)", ("本社", 1))
    print("  - 初期拠点 '本社' を再登録しました。")

    conn.commit()
    print("--- 初期化完了 ---")
    print("※ 管理ユーザー情報は維持されています。")
    
except Exception as e:
    print(f"エラー: {e}")
    conn.rollback()
finally:
    conn.close()
