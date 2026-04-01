from database import engine
from sqlalchemy import text

def cleanup():
    print("--- 不整合なテーブルの清掃を開始します ---")
    with engine.connect() as conn:
        # 新しい機能に関連するテーブルのみを削除し、一から作り直せる状態にします
        conn.execute(text("DROP TABLE IF EXISTS stock_movements"))
        conn.execute(text("DROP TABLE IF EXISTS product_location_stocks"))
        conn.execute(text("DROP TABLE IF EXISTS locations"))
        conn.commit()
    print("--- 100% クリーンな状態にリセットしました！ ---")

if __name__ == "__main__":
    cleanup()
