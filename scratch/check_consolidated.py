from sqlalchemy.orm import Session
from database import SessionLocal
import models

def check_consolidated():
    db = SessionLocal()
    try:
        # 複数の受注を持つ請求書を検索
        invoices = db.query(models.Invoice).all()
        found = False
        for inv in invoices:
            if len(inv.orders) > 1:
                print(f"Found consolidated invoice: ID={inv.id}, Number={inv.invoice_number}, Orders Count={len(inv.orders)}")
                found = True
        
        if not found:
            print("No consolidated invoices found.")
            
            # テスト用に合算請求書を1つ作成してみる（もし必要なら）
            # ここでは検索のみ
    finally:
        db.close()

if __name__ == "__main__":
    check_consolidated()
