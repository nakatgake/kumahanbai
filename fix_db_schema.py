import os
import sys

# プロジェクトルートを追加
sys.path.append(os.getcwd())

from database import engine, SessionLocal
import models
from main import get_password_hash

def fix_and_reset_db():
    print("--- データベース構造の完全再構築（緊急修正）開始 ---")
    
    # 1. 古いテーブルをすべて削除（構造ごと消去）
    models.Base.metadata.drop_all(bind=engine)
    print("  - 古いテーブル構造を完全に削除しました。")
    
    # 2. 最新のモデルに基づいてテーブルを再生成
    models.Base.metadata.create_all(bind=engine)
    print("  - 最新のテーブル構造を生成しました。")
    
    db = SessionLocal()
    try:
        # 3. 初期拠点の登録
        main_loc = models.Location(name="本社", is_default=True)
        db.add(main_loc)
        print("  - 初期拠点 '本社' を登録しました。")
        
        # 4. 管理ユーザーの再登録（ログイン不能を防ぐ）
        # ※中村様のアカウントを再作成します
        admin_user = models.User(
            username="nakamura@connect-web.jp",
            hashed_password=get_password_hash("N687nh4su4"),
            full_name="中村 管理者",
            is_active=True,
            is_admin=True
        )
        db.add(admin_user)
        print("  - 管理ユーザー 'nakamura@connect-web.jp' を再登録しました。")
        
        db.commit()
        print("--- 緊急修正完了 ---")
        print("正常にログインでき、商品登録も可能な状態になりました。")
        
    except Exception as e:
        print(f"エラー: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    fix_and_reset_db()
