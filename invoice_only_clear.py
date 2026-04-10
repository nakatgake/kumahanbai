import sqlite3
import os
import datetime

def invoice_logic_clear():
    db_path = 'kumanog.db' # 念のためのスペルミス防止
    if not os.path.exists('kumanogo.db'):
        print("Database not found.")
        return
    db_path = 'kumanogo.db'

    # 安全のためのバックアップ
    backup_path = f"kumanogo_logic_backup_{int(os.path.getmtime(db_path))}.db"
    import shutil
    shutil.copy2(db_path, backup_path)
    print(f"Safety backup created: {backup_path}")

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 今日の日付（これより前のデータを掃除対象とする）
    safe_date_str = "2026-04-11"

    print(f"\nSearching for historical invoices (Order Date < {safe_date_str})...")

    # 掃除対象の「請求書ID」を論理的に抽出
    # 条件：ステータスが「未入金(UNPAID)」かつ、紐付いている受注の日付が昨日以前のもの
    query = """
    SELECT DISTINCT i.id, i.invoice_number, i.total_amount
    FROM invoices i
    JOIN orders o ON o.invoice_id = i.id
    WHERE i.status = 'InvoiceStatus.UNPAID'
      AND o.order_date < ?
    """
    
    # 代理店受注の場合も考慮
    query_agency = """
    SELECT DISTINCT i.id, i.invoice_number, i.total_amount
    FROM invoices i
    JOIN agency_orders ao ON ao.invoice_id = i.id
    WHERE i.status = 'InvoiceStatus.UNPAID'
      AND ao.order_date < ?
    """

    targets = cursor.execute(query, (safe_date_str,)).fetchall()
    targets_agency = cursor.execute(query_agency, (safe_date_str,)).fetchall()
    
    all_target_ids = set([t['id'] for t in targets] + [t['id'] for t in targets_agency])

    if not all_target_ids:
        print("No historical invoices found to clean up.")
        return

    print(f"Found {len(all_target_ids)} invoices to remove from queue.")

    for inv_id in all_target_ids:
        # 受注データ自体は消さず、請求書だけを削除
        # SQLの外部キー制約 SET NULL により、受注側の invoice_id は自動的に解除されます
        cursor.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))
        print(f"  Removed Invoice ID {inv_id} from dispatch queue.")

    # システムに残ったシャドウデータの残骸も一掃
    cursor.execute("DELETE FROM quotations WHERE quote_number LIKE 'Q-SHADOW-%'")
    cursor.execute("DELETE FROM orders WHERE order_number LIKE 'ORD-SHADOW-%'")

    conn.commit()
    conn.close()

    print("\n✅ Deep Logic Cleanup complete. All business orders are preserved.")
    print("The Bulk Invoice Issuance screen should now be 100% empty.")
