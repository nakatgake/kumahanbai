import datetime
from database import SessionLocal
import models
from main import closing_notification_job

def simulate_closing():
    print("Starting closing simulation...")
    db = SessionLocal()
    try:
        # 今日の日付を「締め日」に該当する日に擬似的に設定してバッチを実行
        # (実際には closing_notification_job 内で today = date.today() しているので、
        #  テスト用に顧客の締め日を今日に合わせる)
        
        today = datetime.date.today()
        test_customer = db.query(models.Customer).filter_by(company="Test Co").first()
        if not test_customer:
            print("Test customer not found. Skipping simulation.")
            return

        print(f"Setting customer {test_customer.id} closing day to {today.day}")
        test_customer.closing_day = today.day
        db.commit()
        
        # バッチ実行
        print("Running closing_notification_job...")
        closing_notification_job()
        
        # 生成された請求書の確認
        latest_inv = db.query(models.Invoice).filter_by(customer_id=test_customer.id).order_by(models.Invoice.id.desc()).first()
        if latest_inv:
            print(f"Verified: Invoice {latest_inv.invoice_number} created.")
            print(f"Total Amount: {latest_inv.total_amount}")
            print(f"Due Date: {latest_inv.due_date}")
            print(f"Memo: {latest_inv.memo}")
        else:
            print("No invoice was generated. (Check if there were orders >= 2026-04-11)")

    except Exception as e:
        print(f"Simulation failed: {e}")
    finally:
        db.close()

if __name__ == "__main__":
    simulate_closing()
