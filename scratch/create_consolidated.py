from sqlalchemy.orm import Session
from database import SessionLocal
import models
import datetime

def create_test_consolidated():
    db = SessionLocal()
    try:
        # 受注を2つ持っている顧客を探す
        customers = db.query(models.Customer).all()
        target_cust = None
        orders_to_combine = []
        
        for cust in customers:
            # 紐付け可能な受注を探す
            pending_orders = db.query(models.Order).join(models.Quotation).filter(
                models.Quotation.customer_id == cust.id
            ).all()
            if len(pending_orders) >= 2:
                target_cust = cust
                orders_to_combine = pending_orders[:2]
                break
        
        if not target_cust:
            print("Could not find a customer with at least 2 orders. Creating dummy orders...")
            # もし見つからなければ適当な顧客に受注を2つ作る
            target_cust = db.query(models.Customer).first()
            if not target_cust:
                print("No customers found in DB.")
                return
            
            for i in range(2):
                q = models.Quotation(customer_id=target_cust.id, quote_number=f"Q-TEST-{datetime.datetime.now().timestamp()}-{i}", total_amount=1000*(i+1), status=models.QuoteStatus.ORDERED)
                db.add(q)
                db.flush()
                o = models.Order(quotation_id=q.id, order_number=f"ORD-TEST-{datetime.datetime.now().timestamp()}-{i}", total_amount=1000*(i+1), status=models.OrderStatus.SHIPPED)
                db.add(o)
                db.flush()
                orders_to_combine.append(o)

        # 既存の紐付けを解除（クリーンアップ）
        for o in orders_to_combine:
            o.invoice_id = None
        
        # 合算請求書の作成
        inv_num = f"INV-CONSOL-{datetime.datetime.now().strftime('%m%d%H%M')}"
        invoice = models.Invoice(
            customer_id=target_cust.id,
            invoice_number=inv_num,
            issue_date=datetime.datetime.now(),
            due_date=datetime.datetime.now() + datetime.timedelta(days=30),
            total_amount=sum(o.total_amount for o in orders_to_combine),
            status=models.InvoiceStatus.UNPAID,
            delivery_status="UNSENT",
            memo="テスト用合算請求書（受注2件統合）"
        )
        db.add(invoice)
        db.flush()
        
        for o in orders_to_combine:
            o.invoice_id = invoice.id
            
        db.commit()
        print(f"Created consolidated invoice: ID={invoice.id}, Number={inv_num} for Customer: {target_cust.company or target_cust.name}")
        print(f"Orders linked: {[o.order_number for o in orders_to_combine]}")
        
    finally:
        db.close()

if __name__ == "__main__":
    create_test_consolidated()
