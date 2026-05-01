import sqlite3
from pathlib import Path


DB_PATH = Path("kumanogo.db")


def rows(cursor, sql, params=()):
    return [dict(row) for row in cursor.execute(sql, params)]


def main():
    if not DB_PATH.exists():
        raise SystemExit(f"database not found: {DB_PATH}")

    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    blocking_checks = {
        "duplicate_invoice_numbers": rows(
            cur,
            """
            SELECT invoice_number, COUNT(*) AS count
            FROM invoices
            GROUP BY invoice_number
            HAVING COUNT(*) > 1
            """,
        ),
        "billable_standard_orders_without_invoice": rows(
            cur,
            """
            SELECT o.id, o.order_number, o.status, o.total_amount
            FROM orders o
            WHERE o.invoice_id IS NULL
              AND o.status IN ('SHIPPED', 'COMPLETED')
            ORDER BY o.id
            """,
        ),
        "billable_agency_orders_without_invoice": rows(
            cur,
            """
            SELECT ao.id, ao.order_number, ao.status, ao.total_amount
            FROM agency_orders ao
            WHERE ao.invoice_id IS NULL
              AND ao.status = '処理済み'
              AND COALESCE(ao.converted_order_id, 0) = 0
            ORDER BY ao.id
            """,
        ),
        "orphan_invoices": rows(
            cur,
            """
            SELECT i.id, i.invoice_number, i.customer_id, i.total_amount, i.status, i.delivery_status
            FROM invoices i
            LEFT JOIN orders o ON o.invoice_id = i.id
            LEFT JOIN agency_orders ao ON ao.invoice_id = i.id
            WHERE o.id IS NULL AND ao.id IS NULL
              AND i.status = 'UNPAID'
              AND COALESCE(i.delivery_status, 'UNSENT') = 'UNSENT'
            ORDER BY i.id
            """,
        ),
        "invoice_total_mismatches": rows(
            cur,
            """
            SELECT i.id, i.invoice_number, i.total_amount,
                   COALESCE((SELECT SUM(o.total_amount) FROM orders o WHERE o.invoice_id = i.id), 0)
                 + COALESCE((SELECT SUM(ao.total_amount) FROM agency_orders ao
                             WHERE ao.invoice_id = i.id AND COALESCE(ao.converted_order_id, 0) = 0), 0)
                   AS linked_total
            FROM invoices i
            WHERE ABS(COALESCE(i.total_amount, 0) - (
                   COALESCE((SELECT SUM(o.total_amount) FROM orders o WHERE o.invoice_id = i.id), 0)
                 + COALESCE((SELECT SUM(ao.total_amount) FROM agency_orders ao
                             WHERE ao.invoice_id = i.id AND COALESCE(ao.converted_order_id, 0) = 0), 0)
            )) > 0.01
              AND i.status = 'UNPAID'
              AND COALESCE(i.delivery_status, 'UNSENT') = 'UNSENT'
            ORDER BY i.id
            """,
        ),
        "converted_agency_orders_missing_link": rows(
            cur,
            """
            SELECT ao.id, ao.order_number, ao.status
            FROM agency_orders ao
            WHERE ao.status = '処理済み'
              AND COALESCE(ao.converted_order_id, 0) = 0
              AND EXISTS (
                  SELECT 1
                  FROM orders o
                  WHERE o.order_number LIKE 'ORD-AG-' || ao.id || '-%'
              )
            ORDER BY ao.id
            """,
        ),
    }

    review_checks = {
        "locked_orphan_invoices_for_manual_review": rows(
            cur,
            """
            SELECT i.id, i.invoice_number, i.customer_id, i.total_amount, i.status, i.delivery_status
            FROM invoices i
            LEFT JOIN orders o ON o.invoice_id = i.id
            LEFT JOIN agency_orders ao ON ao.invoice_id = i.id
            WHERE o.id IS NULL AND ao.id IS NULL
              AND NOT (i.status = 'UNPAID' AND COALESCE(i.delivery_status, 'UNSENT') = 'UNSENT')
            ORDER BY i.id
            """,
        )
    }

    has_issue = False
    for name, result in blocking_checks.items():
        print(f"\n[{name}] {len(result)}")
        for row in result[:50]:
            has_issue = True
            print(row)
        if len(result) > 50:
            print(f"... and {len(result) - 50} more")

    for name, result in review_checks.items():
        print(f"\n[{name}] {len(result)}")
        for row in result[:50]:
            print(row)
        if len(result) > 50:
            print(f"... and {len(result) - 50} more")

    if has_issue:
        raise SystemExit(1)

    print("\nBilling integrity check passed.")


if __name__ == "__main__":
    main()
