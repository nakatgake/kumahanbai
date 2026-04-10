from database import SessionLocal
import models

def print_all_invoices():
    db = SessionLocal()
    try:
        invoices = db.query(models.Invoice).all()
        print(f"Total invoices: {len(invoices)}")
        for inv in invoices:
            print(f"- {inv.invoice_number} | ID: {inv.id} | Status: {inv.status.name if inv.status else 'None'} | Delivery: {inv.delivery_status} | Memo: {inv.memo}")
    except Exception as e:
        import traceback
        traceback.print_exc()
    finally:
        db.close()

if __name__ == "__main__":
    print_all_invoices()
