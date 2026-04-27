from sqlalchemy.orm import Session
from database import SessionLocal
import models
import datetime

def create_huge_consolidated():
    db = SessionLocal()
    try:
        # 商品を取得
        product = db.query(models.Product).first()
        if not product:
            product = models.Product(code="HUGE", name="テスト商品（大量）", unit_price=100)
            db.add(product)
            db.flush()

        target_cust = db.query(models.Customer).filter_by(company="Test Co").first()
        
        # 3つの受注、各10行 = 合計30行+ヘッダー3行 = 33行
        orders_to_combine = []
        for i in range(3):
            q = models.Quotation(
                customer_id=target_cust.id, 
                quote_number=f"Q-HUGE-{datetime.datetime.now().timestamp()}-{i}", 
                total_amount=0, 
                status=models.QuoteStatus.ORDERED,
                issue_date=datetime.date.today()
            )
            db.add(q)
            db.flush()
            
            for j in range(10):
                item = models.QuotationItem(
                    quotation_id=q.id,
                    product_id=product.id,
                    description=f"{product.name} (受注{i}-明細{j})",
                    quantity=1,
                    unit_price=product.unit_price,
                    subtotal=product.unit_price
                )
                db.add(item)
                q.total_amount += item.subtotal
            db.flush()
            
            o = models.Order(
                quotation_id=q.id, 
                order_number=f"ORD-HUGE-{datetime.datetime.now().timestamp()}-{i}", 
                order_date=datetime.datetime.now(),
                total_amount=q.total_amount, 
                status=models.OrderStatus.SHIPPED
            )
            db.add(o)
            db.flush()
            orders_to_combine.append(o)

        inv_num = f"INV-HUGE-{datetime.datetime.now().strftime('%m%d%H%M')}"
        invoice = models.Invoice(
            customer_id=target_cust.id,
            invoice_number=inv_num,
            issue_date=datetime.datetime.now(),
            due_date=datetime.datetime.now() + datetime.timedelta(days=30),
            total_amount=sum(o.total_amount for o in orders_to_combine),
            status=models.InvoiceStatus.UNPAID,
            delivery_status="UNSENT",
            memo="大量明細テスト（30行以上）"
        )
        db.add(invoice)
        db.flush()
        
        for o in orders_to_combine:
            o.invoice_id = invoice.id
            
        db.commit()
        print(f"Created huge consolidated invoice: ID={invoice.id}, Number={inv_num}")
        
    finally:
        db.close()

if __name__ == "__main__":
    create_huge_consolidated()
