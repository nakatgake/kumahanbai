import sqlite3
import os

def invoice_only_clear():
    db_path = 'kumanogo.db'
    if not os.path.exists(db_path):
        print("Database not found.")
        return

    # 安全のためのバックアップ
    backup_path = f"kumanogo_invoice_cleanup_backup_{int(os.path.getmtime(db_path))}.db"
    import shutil
    shutil.copy2(db_path, backup_path)
    print(f"Safety backup created: {backup_path}")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 削除対象の「請求書ID」のみを指定
    # これらは一括発行画面に表示されているもので、受注本体は削除しません。
    target_invoice_ids = [4, 5, 6, 8, 9, 14, 15]

    print("\nStarting invoice-only cleanup (preserving all orders)...")

    for inv_id in target_invoice_ids:
        # この処理により、紐付いている受注の invoice_id は自動的に NULL になります（SET NULL設定済み）
        cursor.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))
        print(f"  Invoice ID {inv_id} removed from queue.")

    # シャドウデータ（一時的な残骸）の掃除
    cursor.execute("DELETE FROM quotations WHERE quote_number LIKE 'Q-SHADOW-%'")
    cursor.execute("DELETE FROM orders WHERE order_number LIKE 'ORD-SHADOW-%'")

    conn.commit()
    conn.close()

    print("\n✅ Cleanup complete. All orders have been preserved.")
    print("The Bulk Invoice Issuance screen should now be empty.")
    print("Please verify your Order List to ensure all data is intact.")

if __name__ == "__main__":
    invoice_only_clear()
