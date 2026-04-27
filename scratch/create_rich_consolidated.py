from sqlalchemy.orm import Session
from database import SessionLocal
import models
import datetime

def create_rich_consolidated():
    db = SessionLocal()
    try:
        # 商品を取得
        product1 = db.query(models.Product).filter(models.Product.name.like("%熊ノ護%")).first()
        product2 = db.query(models.Product).filter(models.Product.name.like("%スプレー%")).first()
        
        if not product1:
            product1 = models.Product(code="TEST01", name="テスト商品A", unit_price=1000)
            db.add(product1)
        if not product2:
            product2 = models.Product(code="TEST02", name="テスト商品B", unit_price=2000)
            db.add(product2)
        db.flush()

        # 既存のテスト用顧客
        target_cust = db.query(models.Customer).filter_by(company="Test Co").first()
        if not target_cust:
            target_cust = models.Customer(company="Test Co", name="Taro", address="秋田市", email="test@example.com")
            db.add(target_cust)
            db.flush()

        orders_to_combine = []
        for i in range(2):
            q = models.Quotation(
                customer_id=target_cust.id, 
                quote_number=f"Q-RICH-{datetime.datetime.now().timestamp()}-{i}", 
                total_amount=0, 
                status=models.QuoteStatus.ORDERED,
                issue_date=datetime.date.today()
            )
            db.add(q)
            db.flush()
            
            # 明細を追加
            item = models.QuotationItem(
                quotation_id=q.id,
                product_id=product1.id if i == 0 else product2.id,
                description=product1.name if i == 0 else product2.name,
                quantity=2 + i,
                unit_price=product1.unit_price if i == 0 else product2.unit_price,
                subtotal=(2 + i) * (product1.unit_price if i == 0 else product2.unit_price)
            )
            db.add(item)
            q.total_amount = item.subtotal
            db.flush()
            
            o = models.Order(
                quotation_id=q.id, 
                order_number=f"ORD-RICH-{datetime.datetime.now().timestamp()}-{i}", 
                order_date=datetime.datetime.now(),
                total_amount=q.total_amount, 
                status=models.OrderStatus.SHIPPED
            )
            db.add(o)
            db.flush()
            orders_to_combine.append(o)

        # 合算請求書の作成
        inv_num = f"INV-RICH-{datetime.datetime.now().strftime('%m%d%H%M')}"
        invoice = models.Invoice(
            customer_id=target_cust.id,
            invoice_number=inv_num,
            issue_date=datetime.datetime.now(),
            due_date=datetime.datetime.now() + datetime.timedelta(days=30),
            total_amount=sum(o.total_amount for o in orders_to_combine),
            status=models.InvoiceStatus.UNPAID,
            delivery_status="UNSENT",
            memo="明細付き合算請求書テスト"
        )
        db.add(invoice)
        db.flush()
        
        for o in orders_to_combine:
            o.invoice_id = invoice.id
            
        db.commit()
        print(f"Created rich consolidated invoice: ID={invoice.id}, Number={inv_num}")
        
    finally:
        db.close()

if __name__ == "__main__":
    create_rich_consolidated()
