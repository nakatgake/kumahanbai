from database import SessionLocal
import models

def seed_settings():
    db = SessionLocal()
    try:
        settings = {
            "tax_rate": "0.1",
            "agency_min_order_amount": "10000",
            "shipping_fee_free_threshold": "30000",
            "shipping_fee": "1200",
            "notification_email": "info@kumanomorikaken.co.jp"
        }
        
        for key, value in settings.items():
            existing = db.query(models.SystemSetting).filter_by(key=key).first()
            if not existing:
                print(f"Seeding setting: {key} = {value}")
                db.add(models.SystemSetting(key=key, value=value))
            else:
                print(f"Setting {key} already exists: {existing.value}")
        
        db.commit()
        print("Settings seeded successfully.")
    except Exception as e:
        db.rollback()
        print(f"Error seeding settings: {e}")
    finally:
        db.close()

if __name__ == "__main__":
    seed_settings()
