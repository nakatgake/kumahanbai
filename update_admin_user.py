import sys
sys.path.append(r"c:\Users\nakatake\Desktop\Antigravityフォルダ\kumanomori")
from database import SessionLocal
import models
from main import get_password_hash

db = SessionLocal()
try:
    # Get the existing admin user
    admin_user = db.query(models.User).filter(models.User.is_admin == True).first()
    if admin_user:
        admin_user.username = "nakamura@connect-web.jp"
        admin_user.hashed_password = get_password_hash("N687nh4su4")
        db.commit()
        print("Admin user updated successfully.")
    else:
        print("Admin user not found. Is the DB empty?")
finally:
    db.close()
