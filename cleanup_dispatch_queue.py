from database import SessionLocal
import models
import traceback

def cleanup_all_unpaid_invoices():
    db = SessionLocal()
    try:
        unpaid = db.query(models.Invoice).filter(models.Invoice.status == models.InvoiceStatus.UNPAID).all()
        print(f"Found {len(unpaid)} UNPAID invoices to clean up.")
        
        for inv in unpaid:
            print(f"- Processing Invoice {inv.invoice_number}")
            
            # 1. Unlink real Agency Orders
            a_orders = db.query(models.AgencyOrder).filter(models.AgencyOrder.invoice_id == inv.id).all()
            for ao in a_orders:
                print(f"  - Unlinking AgencyOrder: {ao.order_number}")
                ao.invoice_id = None
            
            # 2. Unlink standard orders and delete Shadow orders
            orders = db.query(models.Order).filter(models.Order.invoice_id == inv.id).all()
            for o in orders:
                if o.order_number.startswith('ORD-INV-') or o.order_number.startswith('ORD-AGINV-') or ('自動生成' in (o.memo or '')):
                    print(f"  - Deleting Shadow Order: {o.order_number}")
                    q = o.quotation
                    db.delete(o)
                    if q:
                        print(f"    - Deleting Associated Quotation: {q.quote_number}")
                        db.delete(q)
                else:
                    print(f"  - Unlinking Standard Order: {o.order_number}")
                    o.invoice_id = None
            
            # 3. Delete the invoice itself
            db.delete(inv)
            
        # Optional: delete orphaned shadows just in case
        shadows = db.query(models.Order).filter(models.Order.order_number.like("ORD-SHADOW-%")).all()
        for s in shadows:
            q = s.quotation
            db.delete(s)
            if q:
                db.delete(q)

        db.commit()
        print("Cleanup COMPLETED. All test invoices removed from the dispatch queue.")
    except Exception as e:
        db.rollback()
        traceback.print_exc()
        print(f"Error during cleanup: {str(e)}")
    finally:
        db.close()

if __name__ == "__main__":
    cleanup_all_unpaid_invoices()
