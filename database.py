import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# Railwayのpersistent volumeパス、またはローカル実行時はカレントディレクトリを使用
DATA_DIR = os.environ.get("DATA_DIR", ".")

# ディレクトリが存在しない場合は作成する（重要）
if DATA_DIR != "." and not os.path.exists(DATA_DIR):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception as e:
        print(f"Error creating directory {DATA_DIR}: {e}")

SQLALCHEMY_DATABASE_URL = f"sqlite:///{DATA_DIR}/kumanogo.db"

engine = create_engine(
    SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False}
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
