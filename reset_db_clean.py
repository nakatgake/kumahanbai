import sqlite3
import os
import sys

# プロジェクトルートをパスに追加（models, database をインポートするため）
sys.path.append(os.getcwd())

from database import engine, SessionLocal
import models

def reset_db():
    print("--- データベース全データ初期化（完全版）開始 ---")
    
    # 1. テーブルの作成（存在しない場合のみ作成される）
    models.Base.metadata.create_all(bind=engine)
    print("  - 最新のテーブル構造を生成しました。")
    
    db = SessionLocal()
    try:
        # 管理ユーザー（User）以外の全データを削除
        # 順番に注意（外部キー制約のため）
        tables = [
            models.Notification,
            models.SystemSetting,
            models.AgencyOrderItem,
            models.AgencyOrder,
            models.Invoice,
            models.Order,
            models.ProductStock, # 子
            models.QuotationItem,
            models.Quotation,
            models.Location, # 親（product_stocks の後）
            models.Product,
            models.Customer,
        ]
        
        for table in tables:
            count = db.query(table).delete()
            print(f"  - {table.__tablename__} から {count} 件のデータを削除しました。")
        
        # 初期拠点を登録
        main_loc = models.Location(name="本社", is_default=True)
        db.add(main_loc)
        print("  - 初期拠点 '本社' を登録しました。")
        
        db.commit()
        print("--- 初期化完了 ---")
        print("※ 管理ユーザー情報は維持されています。")
        
    except Exception as e:
        print(f"エラー: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    reset_db()
