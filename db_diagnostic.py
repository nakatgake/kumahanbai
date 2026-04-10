import sqlite3
import datetime

def diagnostic():
    db_path = 'kumanogo.db'
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    print("\n" + "="*80)
    print(" DATABASE DIAGNOSTIC REPORT (NON-DESTRUCTIVE)")
    print("="*80)
    print(f"Executed at: {datetime.datetime.now()}")

    # 1. Invoices
    print("\n--- INVOICES (請求書) ---")
    invoices = cursor.execute("SELECT id, invoice_number, issue_date, total_amount, status FROM invoices ORDER BY issue_date DESC").fetchall()
    print(f"{'ID':<4} | {'Invoice Number':<20} | {'Date':<19} | {'Amount':<10} | {'Status'}")
    print("-" * 80)
    for inv in invoices:
        print(f"{inv['id']:<4} | {inv['invoice_number']:<20} | {inv['issue_date'][:19]:<19} | {inv['total_amount']:<10} | {inv['status']}")

    # 2. Orders (Regular)
    print("\n--- REGULAR ORDERS (通常受注) ---")
    orders = cursor.execute("SELECT id, order_number, order_date, total_amount, status, invoice_id FROM orders ORDER BY order_date DESC").fetchall()
    print(f"{'ID':<4} | {'Order Number':<20} | {'Date':<19} | {'Amount':<10} | {'InvID':<5} | {'Status'}")
    print("-" * 80)
    for o in orders:
        inv_id = o['invoice_id'] if o['invoice_id'] else "None"
        print(f"{o['id']:<4} | {o['order_number']:<20} | {o['order_date'][:19]:<19} | {o['total_amount']:<10} | {inv_id:<5} | {o['status']}")

    # 3. Agency Orders
    print("\n--- AGENCY ORDERS (代理店受注) ---")
    try:
        a_orders = cursor.execute("SELECT id, order_number, order_date, total_amount, status, invoice_id FROM agency_orders ORDER BY order_date DESC").fetchall()
        print(f"{'ID':<4} | {'Order Number':<20} | {'Date':<19} | {'Amount':<10} | {'InvID':<5} | {'Status'}")
        print("-" * 80)
        for ao in a_orders:
            inv_id = ao['invoice_id'] if ao['invoice_id'] else "None"
            print(f"{ao['id']:<4} | {ao['order_number']:<20} | {ao['order_date'][:19]:<19} | {ao['total_amount']:<10} | {inv_id:<5} | {ao['status']}")
    except Exception as e:
        print(f"Agency orders table check skipped: {e}")

    conn.close()
    print("\n" + "="*80)
    print(" END OF REPORT")
    print("="*80 + "\n")

if __name__ == "__main__":
    diagnostic()
