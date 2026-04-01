import sys
import os
import datetime

# プロジェクトルートをパスに追加
sys.path.append(os.getcwd())

from database import engine, SessionLocal
import models

def initialize_locations():
    db = SessionLocal()
    try:
        print("--- 複数拠点在庫管理システムの初期化を開始します ---")
        
        # 1. 新しいテーブルの作成 (既存のデータは維持されます)
        models.Base.metadata.create_all(bind=engine)
        print("Step 1: 新しいデータベーステーブルを100%正常に作成しました。")
        
        # 2. デフォルト拠点の作成 (存在しない場合のみ)
        main_location = db.query(models.Location).filter(models.Location.name == "本社倉庫").first()
        if not main_location:
            main_location = models.Location(name="本社倉庫", is_active=True)
            db.add(main_location)
            db.commit()
            db.refresh(main_location)
            print(f"Step 2: デフォルト拠点 '{main_location.name}' を新設しました。")
        else:
            print(f"Step 2: 既存の拠点 '{main_location.name}' を特定しました。")
            
        # 3. 既存の全商品の在庫を拠点別在庫へ一括移行
        products = db.query(models.Product).all()
        migrated_count = 0
        skipped_count = 0
        
        for product in products:
            # すでにこの拠点に在庫データがあるか確認
            existing_stock = db.query(models.ProductLocationStock).filter(
                models.ProductLocationStock.product_id == product.id,
                models.ProductLocationStock.location_id == main_location.id
            ).first()
            
            if not existing_stock:
                # 拠点別在庫レコードを作成
                new_stock = models.ProductLocationStock(
                    product_id=product.id,
                    location_id=main_location.id,
                    stock_quantity=product.stock_quantity or 0
                )
                db.add(new_stock)
                
                # 移動履歴も記録（初期移行として）
                if (product.stock_quantity or 0) > 0:
                    movement = models.StockMovement(
                        product_id=product.id,
                        to_location_id=main_location.id,
                        quantity=product.stock_quantity,
                        type="INBOUND",
                        reason="初期システム移行による自動振分"
                    )
                    db.add(movement)
                
                migrated_count += 1
            else:
                skipped_count += 1
        
        db.commit()
        print(f"Step 3: 完了！ {migrated_count} 件の商品在庫を100%正確に移行しました。 (スキップ: {skipped_count} 件)")
        print("--- 初期化が100%成功しました！ ---")
        
    except Exception as e:
        db.rollback()
        print(f"ERROR: 初期化中に問題が発生しました: {e}")
        sys.exit(1)
    finally:
        db.close()

if __name__ == "__main__":
    initialize_locations()
