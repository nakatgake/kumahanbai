import sqlite3
import os
import datetime

db_path = 'kumanogo.db'
if not os.path.exists(db_path):
    print(f"Error: {db_path} not found.")
    exit(1)

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    print(f"--- 商品台帳・在庫データの修復開始 ({db_path}) ---")
    
    # 1. locations テーブルの作成
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS locations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name VARCHAR NOT NULL,
        description VARCHAR,
        is_default BOOLEAN DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)
    print("  - 'locations' テーブルを作成しました。")
    
    # 2. product_stocks テーブルの作成
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS product_stocks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL,
        location_id INTEGER NOT NULL,
        quantity INTEGER DEFAULT 0,
        FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE,
        FOREIGN KEY(location_id) REFERENCES locations(id) ON DELETE CASCADE
    )
    """)
    print("  - 'product_stocks' テーブルを作成しました。")
    
    # 3. 初期拠点の投入 (なければ)
    cursor.execute("SELECT COUNT(*) FROM locations")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO locations (name, is_default) VALUES (?, ?)", ("本社", 1))
        location_id = cursor.lastrowid
        print(f"  - デフォルト拠点 '本社' (ID: {location_id}) を登録しました。")
    else:
        cursor.execute("SELECT id FROM locations LIMIT 1")
        location_id = cursor.fetchone()[0]
        print(f"  - 既存拠点 (ID: {location_id}) を使用します。")
        
    # 4. 既存商品の在庫を product_stocks に同期
    cursor.execute("SELECT id, stock_quantity FROM products")
    products = cursor.fetchall()
    
    synced_count = 0
    for pid, qty in products:
        # すでに紐付けがないか確認
        cursor.execute("SELECT COUNT(*) FROM product_stocks WHERE product_id = ? AND location_id = ?", (pid, location_id))
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO product_stocks (product_id, location_id, quantity) VALUES (?, ?, ?)", 
                           (pid, location_id, qty if qty else 0))
            synced_count += 1
            
    print(f"  - {synced_count} 件の商品在庫を拠点に同期しました。")

    conn.commit()
    print("--- 修復完了 ---")
    
except Exception as e:
    print(f"エラー: {e}")
finally:
    conn.close()
