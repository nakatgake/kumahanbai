import sqlite3
import os

def zero_clear():
    db_path = 'kumanogo.db'
    if not os.path.exists(db_path):
        print("Database not found.")
        return

    # バックアップの作成（念のため）
    backup_path = f"kumanogo_before_zero_{int(os.path.getmtime(db_path))}.db"
    import shutil
    shutil.copy2(db_path, backup_path)
    print(f"Safety backup created at: {backup_path}")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. 削除対象の定義（診断レポートに基づく）
    target_invoice_ids = [4, 5, 6, 8, 9, 14, 15]
    target_order_ids = [4, 5, 6, 8, 9, 13, 15]

    print("\nStarting zero-based targeted cleanup...")

    # A. 請求書の削除
    for inv_id in target_invoice_ids:
        print(f"  Deleting Invoice ID {inv_id}...")
        cursor.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))

    # B. 通常受注の削除（および関連する見積もり）
    for order_id in target_order_ids:
        print(f"  Deleting Order ID {order_id}...")
        # 紐付く見積もりIDを取得
        row = cursor.execute("SELECT quotation_id FROM orders WHERE id = ?", (order_id,)).fetchone()
        if row and row[0]:
            quote_id = row[0]
            cursor.execute("DELETE FROM quotations WHERE id = ?", (quote_id,))
            print(f"    - Linked Quotation ID {quote_id} deleted.")
        cursor.execute("DELETE FROM orders WHERE id = ?", (order_id,))

    # C. 浮いているシャドウデータの掃除
    cursor.execute("DELETE FROM quotations WHERE quote_number LIKE 'Q-SHADOW-%'")
    cursor.execute("DELETE FROM orders WHERE order_number LIKE 'ORD-SHADOW-%'")

    # D. 紐付けの修正（ゴミ請求書を参照している受注が他にもあれば解除のみ）
    cursor.execute(f"UPDATE orders SET invoice_id = NULL WHERE invoice_id IN ({','.join(['?']*len(target_invoice_ids))})", target_invoice_ids)
    cursor.execute(f"UPDATE agency_orders SET invoice_id = NULL WHERE invoice_id IN ({','.join(['?']*len(target_invoice_ids))})", target_invoice_ids)

    conn.commit()
    conn.close()
    print("\n✅ Targeted cleanup complete.")
    print("Please refresh /admin/diagnostic and verify that IDs 16, 17, 18, 19 remain intact.")

if __name__ == "__main__":
    zero_clear()
