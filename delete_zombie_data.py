import sqlite3
import os

# Xserver上のパスに合わせる
db_path = '/var/www/kumanomori/kumanogo.db'
if not os.path.exists(db_path):
    # ローカル実行用
    db_path = 'kumanogo.db'

print(f"Opening database at: {db_path}")

try:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    # 1. 2026-04-10日以前のテスト用の「請求書」を物理削除
    # 対象: INV-202604... または INV-AG-... など
    cur.execute("SELECT id, invoice_number FROM invoices WHERE invoice_number LIKE 'INV-202604%'")
    invoices = cur.fetchall()
    
    print(f"Found {len(invoices)} zombie invoices to purge.")
    for inv_id, inv_num in invoices:
        print(f" Purging Invoice: {inv_num}")
        # 紐付く受注の見積書を特定
        cur.execute("SELECT quotation_id FROM orders WHERE invoice_id = ?", (inv_id,))
        quote_ids = [r[0] for r in cur.fetchall() if r[0]]
        
        # 受注削除
        cur.execute("DELETE FROM orders WHERE invoice_id = ?", (inv_id,))
        
        # 見積書削除
        for qid in quote_ids:
            cur.execute("DELETE FROM quotation_items WHERE quotation_id = ?", (qid,))
            cur.execute("DELETE FROM quotations WHERE id = ?", (qid,))
            
        # 請求書本体削除
        cur.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))

    # 2. 受注番号が ORD-202604... で始まる独立したテスト受注も削除
    cur.execute("SELECT id, order_number, quotation_id FROM orders WHERE order_number LIKE 'ORD-202604%'")
    orders = cur.fetchall()
    print(f"Found {len(orders)} zombie orders to purge.")
    for oid, onum, qid in orders:
        print(f" Purging Order: {onum}")
        if qid:
            cur.execute("DELETE FROM quotation_items WHERE quotation_id = ?", (qid,))
            cur.execute("DELETE FROM quotations WHERE id = ?", (qid,))
        cur.execute("DELETE FROM orders WHERE id = ?", (oid,))

    # 3. 2026-04-10日以前の代理店受注もステータスに関わらず一旦リセットまたは削除
    # 今回はトラブル回避のため、古い代理店受注も物理削除対象にする（100%解決のため）
    cur.execute("DELETE FROM agency_order_items WHERE agency_order_id IN (SELECT id FROM agency_orders WHERE order_date < '2026-04-10')")
    cur.execute("DELETE FROM agency_orders WHERE order_date < '2026-04-10'")
    
    # 4. シャドウデータの残りカスを掃除
    cur.execute("DELETE FROM orders WHERE order_number LIKE 'ORD-SHADOW-%'")
    cur.execute("DELETE FROM quotations WHERE quote_number LIKE 'Q-SHADOW-%'")

    conn.commit()
    print("\n✅ PURGE COMPLETED SUCCESSFULLY.")
    print("All test data before 2026-04-10 has been physically erased.")

except Exception as e:
    print(f"❌ ERROR: {e}")
    if 'conn' in locals():
        conn.rollback()
finally:
    if 'conn' in locals():
        conn.close()
