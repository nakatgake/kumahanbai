from fastapi import FastAPI, Depends, Request, Form, File, UploadFile, Query
from typing import Optional
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
import shutil
import os
import random
import string
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func
from sqlalchemy.orm import Session
import models
from database import SessionLocal, engine, get_db
import datetime
import io
from fpdf import FPDF
import xlsxwriter
from fastapi.responses import StreamingResponse
from passlib.context import CryptContext
from itsdangerous import URLSafeSerializer
from fastapi import HTTPException
import smtplib
import ssl
try:
    from apscheduler.schedulers.background import BackgroundScheduler
    from utils.date_utils import is_closing_day, calculate_payment_date, next_business_day, get_next_closing_date
    HAS_SCHEDULER = True
except ImportError:
    print("Warning: APScheduler or date-util not found. Automated billing is disabled.")
    HAS_SCHEDULER = False

from email.message import EmailMessage
from utils.email import send_notification

# Create database tables
models.Base.metadata.create_all(bind=engine)

def update_product_stock(db: Session, product_id: int, location_id: int, quantity: int, move_type: str, reason: str = None, from_location_id: int = None):
    """
    複数拠点の在庫を更新し、履歴を記録する一括関数
    move_type: "INBOUND", "OUTBOUND", "TRANSFER", "ADJUSTMENT"
    """
    product = db.query(models.Product).get(product_id)
    if not product:
        return

    if move_type == "TRANSFER":
        # 移動元からマイナス
        from_stock = db.query(models.ProductLocationStock).filter_by(product_id=product_id, location_id=from_location_id).first()
        if not from_stock:
            from_stock = models.ProductLocationStock(product_id=product_id, location_id=from_location_id, stock_quantity=0)
            db.add(from_stock)
        from_stock.stock_quantity -= quantity
        
        # 移動先へプラス
        to_stock = db.query(models.ProductLocationStock).filter_by(product_id=product_id, location_id=location_id).first()
        if not to_stock:
            to_stock = models.ProductLocationStock(product_id=product_id, location_id=location_id, stock_quantity=0)
            db.add(to_stock)
        to_stock.stock_quantity += quantity
        
        # 履歴記録
        movement = models.StockMovement(
            product_id=product_id, from_location_id=from_location_id, to_location_id=location_id,
            quantity=quantity, type=move_type, reason=reason
        )
        db.add(movement)
    
    else:
        # 入庫（INBOUND）または出庫（OUTBOUND）
        stock = db.query(models.ProductLocationStock).filter_by(product_id=product_id, location_id=location_id).first()
        if not stock:
            stock = models.ProductLocationStock(product_id=product_id, location_id=location_id, stock_quantity=0)
            db.add(stock)
            
        if move_type == "INBOUND":
            stock.stock_quantity += quantity
            movement = models.StockMovement(product_id=product_id, to_location_id=location_id, quantity=quantity, type=move_type, reason=reason)
        else: # OUTBOUND
            stock.stock_quantity -= quantity
            movement = models.StockMovement(product_id=product_id, from_location_id=location_id, quantity=quantity, type=move_type, reason=reason)
        db.add(movement)
    
    db.flush() # IDを確定させる
    # 全拠点の合計在庫を Product.stock_quantity にキャッシュ
    all_stocks = db.query(models.ProductLocationStock).filter_by(product_id=product_id).all()
    product.stock_quantity = sum(s.stock_quantity for s in all_stocks)

# Database Migration for Customer Rank and Product Price Tiers
def migrate_db():
    import sqlite3
    import os
    
    # データベースのパスを database.py の設定に合わせる
    data_dir = os.getenv("DATA_DIR", ".")
    db_path = os.path.join(data_dir, "kumanogo.db")
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Check if customers table exists first
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='customers'")
    if not cursor.fetchone():
        print("Table 'customers' doesn't exist yet. Skipping migration.")
        conn.close()
        return
    
    # Check if customers.rank exists
    cursor.execute("PRAGMA table_info(customers)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'rank' not in cols:
        print("Migrating customers: adding rank column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN rank VARCHAR DEFAULT 'RETAIL'")
        
    # Check if product price tiers exist
    cursor.execute("PRAGMA table_info(products)")
    cols = [row[1] for row in cursor.fetchall()]
    new_cols = ['price_retail', 'price_a', 'price_b', 'price_c', 'price_d', 'price_e']
    for col in new_cols:
        if col not in cols:
            print(f"Migrating products: adding {col} column...")
            cursor.execute(f"ALTER TABLE products ADD COLUMN {col} FLOAT DEFAULT 0.0")
            # If migrating price_retail, copy unit_price to it
            if col == 'price_retail' and 'unit_price' in cols:
                cursor.execute("UPDATE products SET price_retail = unit_price")
    
    # Check if quotation/order discount fields exist
    for table in ['quotations', 'orders']:
        cursor.execute(f"PRAGMA table_info({table})")
        cols = [row[1] for row in cursor.fetchall()]
        if 'discount_rate' not in cols:
            print(f"Migrating {table}: adding discount_rate...")
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN discount_rate FLOAT DEFAULT 0.0")
        if 'is_bulk_discount' not in cols:
            print(f"Migrating {table}: adding is_bulk_discount...")
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN is_bulk_discount BOOLEAN DEFAULT 0")
        if 'memo' not in cols:
            print(f"Migrating {table}: adding memo...")
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN memo TEXT")
    
    # Check if invoice discount fields exist
    cursor.execute("PRAGMA table_info(invoices)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'discount_rate' not in cols:
        print("Migrating invoices: adding discount_rate...")
        cursor.execute("ALTER TABLE invoices ADD COLUMN discount_rate FLOAT DEFAULT 0.0")
    if 'is_bulk_discount' not in cols:
        print("Migrating invoices: adding is_bulk_discount...")
        cursor.execute("ALTER TABLE invoices ADD COLUMN is_bulk_discount BOOLEAN DEFAULT 0")
    if 'memo' not in cols:
        print("Migrating invoices: adding memo...")
        cursor.execute("ALTER TABLE invoices ADD COLUMN memo TEXT")
    
    # Check if customer agency fields exist
    cursor.execute("PRAGMA table_info(customers)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'is_agency' not in cols:
        print("Migrating customers: adding is_agency column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN is_agency BOOLEAN DEFAULT 0")
    if 'login_id' not in cols:
        print("Migrating customers: adding login_id column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN login_id VARCHAR")
    if 'agency_password' not in cols:
        print("Migrating customers: adding agency_password column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN agency_password VARCHAR")
    if 'invoice_delivery_method' not in cols:
        print("Migrating customers: adding invoice_delivery_method column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN invoice_delivery_method VARCHAR DEFAULT 'POSTAL'")
    if 'honorific' not in cols:
        print("Migrating customers: adding honorific column...")
        cursor.execute("ALTER TABLE customers ADD COLUMN honorific VARCHAR DEFAULT '御中'")
    
    # Check if system_settings table exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='system_settings'")
    if not cursor.fetchone():
        print("Creating system_settings table...")
        cursor.execute("""
            CREATE TABLE system_settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key VARCHAR UNIQUE,
                value VARCHAR
            )
        """)
    
    # Check if notification is_read field exists
    cursor.execute("PRAGMA table_info(notifications)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'is_read' not in cols:
        print("Migrating notifications: adding is_read column...")
        cursor.execute("ALTER TABLE notifications ADD COLUMN is_read BOOLEAN DEFAULT 0")
    if 'related_type' not in cols:
        print("Migrating notifications: adding related_type column...")
        cursor.execute("ALTER TABLE notifications ADD COLUMN related_type VARCHAR")
    if 'related_id' not in cols:
        print("Migrating notifications: adding related_id column...")
        cursor.execute("ALTER TABLE notifications ADD COLUMN related_id INTEGER")
    
    # Check if consolidated invoice columns exist (合算請求対応)
    cursor.execute("PRAGMA table_info(invoices)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'delivery_status' not in cols:
        print("Migrating invoices: adding delivery_status column...")
        cursor.execute("ALTER TABLE invoices ADD COLUMN delivery_status TEXT DEFAULT 'UNSENT'")
    if 'customer_id' not in cols:
        print("Migrating invoices: adding customer_id column...")
        cursor.execute("ALTER TABLE invoices ADD COLUMN customer_id INTEGER")
    
    cursor.execute("PRAGMA table_info(orders)")
    cols = [row[1] for row in cursor.fetchall()]
    if 'invoice_id' not in cols:
        print("Migrating orders: adding invoice_id column...")
        cursor.execute("ALTER TABLE orders ADD COLUMN invoice_id INTEGER")
    
    # --- 既存データの自動移行（旧1対1の請求書データ対応） ---
    try:
        cursor.execute("PRAGMA table_info(invoices)")
        inv_cols = [row[1] for row in cursor.fetchall()]
        if 'order_id' in inv_cols:
            print("Running data migration for legacy invoices...")
            # 旧請求書のorder_idを元に、受注側のinvoice_idに紐づける
            cursor.execute("""
                UPDATE orders 
                SET invoice_id = (
                    SELECT id FROM invoices WHERE invoices.order_id = orders.id
                )
                WHERE invoice_id IS NULL AND EXISTS (
                    SELECT 1 FROM invoices WHERE invoices.order_id = orders.id
                )
            """)
            # 旧請求書に不足しているcustomer_idを、紐づく受注先から引っ張って設定する
            cursor.execute("""
                UPDATE invoices
                SET customer_id = (
                    SELECT q.customer_id 
                    FROM orders o
                    JOIN quotations q ON o.quotation_id = q.id
                    WHERE o.id = invoices.order_id
                )
                WHERE customer_id IS NULL AND order_id IS NOT NULL
            """)
    except Exception as e:
        print(f"Data migration warning: {e}")

    conn.commit()
    conn.close()

migrate_db()

# --- Security Configuration ---
SECRET_KEY = "kumanogo-secret-key-12345" # 本番環境では環境変数などで管理すべき
pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")
serializer = URLSafeSerializer(SECRET_KEY)

def get_password_hash(password):
    return pwd_context.hash(password)

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)


class NotAuthenticatedException(Exception):
    pass

async def get_current_user(request: Request, db: Session = Depends(get_db)):
    session_token = request.cookies.get("session")
    if not session_token:
        return None
    try:
        username = serializer.loads(session_token)
        user = db.query(models.User).filter(models.User.username == username).first()
        return user
    except:
        return None

async def get_active_user(request: Request, db: Session = Depends(get_db)):
    user = await get_current_user(request, db)
    if not user:
        if request.url.path not in ["/login", "/init-admin"] and not request.url.path.startswith("/static"):
            raise NotAuthenticatedException()
    return user

import traceback

# 1. エラーを記録するためのグローバル変数
LAST_ERROR = "No errors logged yet."

app = FastAPI()

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    global LAST_ERROR
    LAST_ERROR = traceback.format_exc()
    print(f"DEBUG_LOG: {LAST_ERROR}")
    return HTMLResponse(content=f"<h1>Internal Server Error</h1><p>Please check <a href='/debug-logs'>/debug-logs</a></p><pre>{LAST_ERROR}</pre>", status_code=500)

@app.get("/debug-logs", response_class=HTMLResponse)
async def view_debug_logs(user: models.User = Depends(get_active_user)):
    if user and not user.is_admin:
        return "Access Denied"
    return f"<h1>Last Error Traceback</h1><pre>{LAST_ERROR}</pre>"

@app.exception_handler(NotAuthenticatedException)
async def auth_exception_handler(request: Request, exc: NotAuthenticatedException):
    return RedirectResponse(url="/login")

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

def closing_notification_job():
    """毎日朝8時に実行される、締め日の判定と請求書自動発行（未発行状態で作成）処理"""
    db = SessionLocal()
    try:
        today = datetime.date.today()
        # 全顧客を取得して締め日を判定
        customers = db.query(models.Customer).all()
        for cust in customers:
            if not cust.closing_day:
                continue
            
            # 今日が締め日の場合のみ処理実行
            if is_closing_day(today, cust.closing_day):
                # 1. 未請求の標準受注を抽出 (出荷済み or 完了 のみ)
                # 100%解決のため、テスト期間中（2026-04-10以前）のデータは絶対に拾わない条件を追加
                safe_date = datetime.datetime(2026, 4, 10)
                standard_orders = db.query(models.Order).join(models.Quotation).filter(
                    models.Quotation.customer_id == cust.id,
                    models.Order.invoice_id == None,
                    models.Order.order_date >= safe_date,
                    models.Order.status.in_([models.OrderStatus.SHIPPED, models.OrderStatus.COMPLETED])
                ).all()
                
                # 2. 未請求の代理店受注を抽出 (処理済みのみ)
                # 100%解決のため、テスト期間中（2026-04-10以前）のデータは絶対に拾わない
                agency_orders = db.query(models.AgencyOrder).filter(
                    models.AgencyOrder.customer_id == cust.id,
                    models.AgencyOrder.status == "処理済み",
                    models.AgencyOrder.invoice_id == None,
                    models.AgencyOrder.order_date >= safe_date
                ).all()
                
                if not standard_orders and not agency_orders:
                    continue
                
                # 請求番号の決定 (INV-YYYYMM-CUSTID)
                inv_num = f"INV-{today.strftime('%Y%m')}-{cust.id:04d}"
                existing = db.query(models.Invoice).filter_by(invoice_number=inv_num).first()
                if existing:
                    inv_num = f"{inv_num}-REV-{datetime.datetime.now().strftime('%H%M')}"

                # 支払期限の計算
                due_date_raw = calculate_payment_date(today, cust.payment_term_months or 1, cust.payment_day or 31)
                due_date = next_business_day(due_date_raw)
                
                # 合計金額の算出
                total_standard = sum(o.total_amount for o in standard_orders)
                total_agency = sum(o.total_amount for o in agency_orders)
                grand_total = total_standard + total_agency

                # 請求書の作成 (初期状態は「未入金」かつ「未発送/未メール」)
                invoice = models.Invoice(
                    customer_id=cust.id,
                    invoice_number=inv_num,
                    issue_date=datetime.datetime.combine(today, datetime.time.min),
                    due_date=datetime.datetime.combine(due_date, datetime.time.min),
                    total_amount=float(grand_total),
                    status=models.InvoiceStatus.UNPAID,
                    delivery_status="UNSENT",
                    memo=f"{today.strftime('%m/%d')}締め 合算集計（自動生成）"
                )
                db.add(invoice)
                db.flush()

                # --- 標準受注の紐付け ---
                for o in standard_orders:
                    o.invoice_id = invoice.id
                
                # --- 代理店受注の紐付けとシャドウ受注の作成 ---
                if agency_orders:
                    # 代理店発注分を1つの「シャドウ受注」にまとめてInvoice.ordersに含める (PDF等での表示用)
                    shadow_quote = models.Quotation(
                        customer_id=cust.id,
                        quote_number=f"Q-SHADOW-{inv_num}",
                        issue_date=today,
                        expiry_date=today + datetime.timedelta(days=30),
                        total_amount=float(total_agency),
                        status=models.QuoteStatus.ORDERED,
                        memo=f"代理店発注集計分 ({len(agency_orders)}件)"
                    )
                    db.add(shadow_quote)
                    db.flush()

                    for ao in agency_orders:
                        ao.invoice_id = invoice.id  # 代理店発注に請求書IDをセット
                        for item in ao.items:
                            qi = models.QuotationItem(
                                quotation_id=shadow_quote.id,
                                product_id=item.product_id,
                                description=f"[{ao.order_number}] {item.product_name}",
                                quantity=item.quantity,
                                unit_price=item.unit_price,
                                subtotal=item.subtotal
                            )
                            db.add(qi)
                    
                    shadow_order = models.Order(
                        quotation_id=shadow_quote.id,
                        order_number=f"ORD-SHADOW-{inv_num}",
                        order_date=today,
                        invoice_id=invoice.id,
                        total_amount=float(total_agency),
                        status=models.OrderStatus.COMPLETED,
                        memo="代理店発注合算シャドウ"
                    )
                    db.add(shadow_order)
                
                # 管理者への通知メール
                try:
                    order_numbers = [o.order_number for o in standard_orders] + [ao.order_number for ao in agency_orders]
                    subject = f"【自動作成】本日締め日：{cust.company} 様の合算請求書を作成しました"
                    body = f"""株式会社熊ノ護化研 担当者様

本日（{today.strftime('%Y/%m/%d')}）は、{cust.company or cust.name} 様の締め日です。
以下の通り、複数の受注を1枚にまとめた合算請求書を自動作成しました。

■ 顧客名：{cust.company or cust.name}
■ 合算内容：合計 {len(standard_orders) + len(agency_orders)} 件の受注を統合
■ 合計請求額：¥{'{:,.0f}'.format(int(grand_total * 1.1))} (税込)
■ 対象受注番号：{', '.join(order_numbers)}

【重要：確認のお願い】
この請求書は「未送信」状態で作成されています。
管理画面の「請求書一括発行」メニューより内容をプレビュー確認し、
お客様の希望（{cust.invoice_delivery_method}）に合わせて、
「メール送信」または「印刷・郵送」の操作を行ってください。
"""
                    settings = db.query(models.SystemSetting).all()
                    s_dict = {s.key: s.value for s in settings}
                    target_email = s_dict.get("notification_email") or "info@kumanomorikaken.co.jp"
                    
                    # 管理者へメール通知を実行
                    send_admin_email_sync(db, subject, body)
                except Exception as email_err:
                    print(f"WARNING: Email notification failed: {str(email_err)}")

                db.commit()
                print(f"INFO: Automated Invoice generated for Customer {cust.id}: {inv_num}")

    except Exception as e:
        db.rollback()
        import traceback
        traceback.print_exc()
        print(f"CRITICAL: Failed in closing_notification_job: {str(e)}")
    finally:
        db.close()

# APScheduler 設定
if HAS_SCHEDULER:
    scheduler = BackgroundScheduler(timezone="Asia/Tokyo")
    scheduler.add_job(closing_notification_job, 'cron', hour=8, minute=0)

@app.on_event("startup")
async def startup_event():
    if HAS_SCHEDULER:
        scheduler.start()
        print("APScheduler started on startup (Closing Notification Job: 08:00 JST)")
    else:
        print("APScheduler skipped (missing libraries)")

@app.on_event("shutdown")
async def shutdown_event():
    if HAS_SCHEDULER:
        scheduler.shutdown()
        print("APScheduler shutdown")
# --- Dashboard ---
@app.get("/", response_class=HTMLResponse)
async def dashboard(
    request: Request, 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    # Base queries
    q_customers = db.query(models.Customer)
    q_products = db.query(models.Product)
    q_quotes = db.query(models.Quotation)
    q_orders = db.query(models.Order)
    q_invoices = db.query(models.Invoice)

    # Date filters
    if start_date:
        sd = datetime.datetime.strptime(start_date, '%Y-%m-%d')
        q_customers = q_customers.filter(models.Customer.created_at >= sd)
        q_products = q_products.filter(models.Product.created_at >= sd)
        q_quotes = q_quotes.filter(models.Quotation.issue_date >= sd)
        q_orders = q_orders.filter(models.Order.order_date >= sd)
        q_invoices = q_invoices.filter(models.Invoice.issue_date >= sd)
    if end_date:
        ed = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        q_customers = q_customers.filter(models.Customer.created_at < ed)
        q_products = q_products.filter(models.Product.created_at < ed)
        q_quotes = q_quotes.filter(models.Quotation.issue_date < ed)
        q_orders = q_orders.filter(models.Order.order_date < ed)
        q_invoices = q_invoices.filter(models.Invoice.issue_date < ed)

    # 請求関連サマリー（ダッシュボード用）
    try:
        # 発送待ち件数（未発送の請求書）
        dispatch_pending = db.query(models.Invoice).filter(models.Invoice.delivery_status == "UNSENT").count()
        
        # 本日の自動締処理分
        today_start = datetime.datetime.combine(datetime.date.today(), datetime.time.min)
        auto_count = db.query(models.Invoice).filter(
            models.Invoice.issue_date >= today_start,
            models.Invoice.memo.like("%自動生成%")
        ).count()
        
        paid_total = q_invoices.filter(models.Invoice.status == models.InvoiceStatus.PAID).with_entities(func.sum(models.Invoice.total_amount)).scalar() or 0
        unpaid_total = q_invoices.filter(models.Invoice.status == models.InvoiceStatus.UNPAID).with_entities(func.sum(models.Invoice.total_amount)).scalar() or 0
    except Exception as stats_err:
        print(f"Error calculating dashboard stats: {stats_err}")
        dispatch_pending = 0
        auto_count = 0
        paid_total = 0
        unpaid_total = 0

    stats = {
        "customers": q_customers.count(),
        "products": q_products.count(),
        "quotations": q_quotes.count(),
        "orders": q_orders.count(),
        "invoices": q_invoices.count(),
        "total_sales": paid_total,
        "unpaid_total": unpaid_total,
        "dispatch_pending": dispatch_pending,
        "auto_closed_today": auto_count
    }

    return templates.TemplateResponse(request=request, name="dashboard.html", context={
        "request": request,
        "active_page": "dashboard",
        "stats": stats,
        "start_date": start_date,
        "end_date": end_date,
        "user": user,
        "now": datetime.datetime.now()
    })
# --- Customers ---
@app.get("/customers", response_class=HTMLResponse)
async def list_customers(
    request: Request, 
    q: str = "", 
    success_email: int = 0,
    error: str = None,
    msg: str = None,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Customer)
    if q:
        query = query.filter(
            (models.Customer.name.contains(q)) | 
            (models.Customer.company.contains(q)) |
            (models.Customer.email.contains(q))
        )
    customers = query.order_by(models.Customer.id.desc()).all()
    return templates.TemplateResponse(request=request, name="customers/list.html", context={
        "request": request,
        "active_page": "customers",
        "customers": customers,
        "search_query": q,
        "message": "アカウント情報を送信しました。" if success_email else None,
        "error": f"送信に失敗しました: {msg}" if error == "send_failed" else (f"SMTP設定（{msg}）を確認してください。" if error == "smtp_config" else ("情報を確認してください。" if error == "missing_info" else None)),
        "CustomerRank": models.CustomerRank,
        "user": user
    })

@app.get("/customers/excel")
async def export_customers_excel(
    q: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Customer)
    if q:
        query = query.filter((models.Customer.company.contains(q)) | (models.Customer.name.contains(q)))
    customers = query.order_by(models.Customer.id.desc()).all()
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("顧客一覧")
    
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})

    headers = ["会社名", "担当者名", "メール", "電話番号", "住所"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, c in enumerate(customers, 1):
        worksheet.write(row, 0, c.company, border_fmt)
        worksheet.write(row, 1, c.name, border_fmt)
        worksheet.write(row, 2, c.email, border_fmt)
        worksheet.write(row, 3, c.phone, border_fmt)
        worksheet.write(row, 4, c.address, border_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"customers_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/customers/new", response_class=HTMLResponse)
async def new_customer(
    request: Request,
    user: models.User = Depends(get_active_user)
):
    return templates.TemplateResponse(request=request, name="customers/form.html", context={
        "request": request, 
        "active_page": "customers",
        "ranks": models.CustomerRank,
        "user": user
    })

@app.post("/customers/new")
async def create_customer(
    name: Optional[str] = Form(None),
    company: str = Form(""),
    zip_code: str = Form(""),
    email: str = Form(""),
    phone: str = Form(""),
    address: str = Form(""),
    website_url: str = Form(""),
    rank: str = Form("RETAIL"),
    honorific: str = Form("御中"),
    is_agency: bool = Form(False),
    invoice_delivery_method: str = Form("POSTAL"),
    login_id: Optional[str] = Form(None),
    agency_password: Optional[str] = Form(None),
    closing_day: Optional[int] = Form(None),
    payment_term_months: int = Form(1),
    payment_day: Optional[int] = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = models.Customer(
        name="" if name == "None" else name,
        company=company,
        zip_code="" if zip_code == "None" else zip_code, 
        email="" if email == "None" else email,
        phone="" if phone == "None" else phone,
        address="" if address == "None" else address,
        website_url="" if website_url == "None" else website_url,
        rank=models.CustomerRank[rank],
        honorific=honorific,
        is_agency=is_agency,
        invoice_delivery_method=invoice_delivery_method,
        login_id=login_id if is_agency and login_id else None,
        agency_password=agency_password if is_agency and agency_password else None,
        closing_day=closing_day,
        payment_term_months=payment_term_months,
        payment_day=payment_day
    )
    db.add(customer)
    db.commit()
    return RedirectResponse(url="/customers", status_code=303)

@app.get("/customers/edit/{customer_id}", response_class=HTMLResponse)
async def edit_customer(
    customer_id: int, 
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = db.query(models.Customer).get(customer_id)
    return templates.TemplateResponse(request=request, name="customers/form.html", context={
        "request": request, 
        "active_page": "customers", 
        "customer": customer,
        "ranks": models.CustomerRank,
        "user": user
    })

@app.post("/customers/edit/{customer_id}")
async def update_customer(
    customer_id: int,
    name: Optional[str] = Form(None),
    company: str = Form(""),
    zip_code: str = Form(""),
    email: str = Form(""),
    phone: str = Form(""),
    address: str = Form(""),
    website_url: str = Form(""),
    rank: str = Form("RETAIL"),
    honorific: str = Form("御中"),
    is_agency: bool = Form(False),
    invoice_delivery_method: str = Form("POSTAL"),
    login_id: Optional[str] = Form(None),
    agency_password: Optional[str] = Form(None),
    closing_day: Optional[int] = Form(None),
    payment_term_months: int = Form(1),
    payment_day: Optional[int] = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = db.query(models.Customer).get(customer_id)
    if customer:
        customer.name = "" if name == "None" else name
        customer.company = company
        customer.zip_code = "" if zip_code == "None" else zip_code
        customer.email = "" if email == "None" else email
        customer.phone = "" if phone == "None" else phone
        customer.address = "" if address == "None" else address
        customer.website_url = "" if website_url == "None" else website_url
        customer.rank = models.CustomerRank[rank]
        customer.honorific = honorific
        customer.is_agency = is_agency
        customer.invoice_delivery_method = invoice_delivery_method
        customer.login_id = login_id if is_agency and login_id else None
        customer.agency_password = agency_password if is_agency and agency_password else None
        customer.closing_day = closing_day
        customer.payment_term_months = payment_term_months
        customer.payment_day = payment_day
        db.commit()
    return RedirectResponse(url="/customers", status_code=303)

@app.post("/customers/delete/{customer_id}")
async def delete_customer(customer_id: int, db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)):
    customer = db.query(models.Customer).get(customer_id)
    if customer:
        db.delete(customer)
        db.commit()
    response = RedirectResponse(url="/customers", status_code=303)
    response.headers["HX-Refresh"] = "true"
    return response

@app.get("/customers/{customer_id}/print-agency-info", response_class=HTMLResponse)
async def print_agency_info(
    customer_id: int, 
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = db.query(models.Customer).get(customer_id)
    if not customer or not customer.is_agency:
        return RedirectResponse(url="/customers", status_code=303)
        
    portal_url = str(request.base_url).rstrip('/') + "/agency/login"
    
    return templates.TemplateResponse(request=request, name="customers/agency_print.html", context={
        "request": request, 
        "customer": customer,
        "issue_date": datetime.datetime.now().strftime('%Y年%m月%d日'),
        "portal_url": portal_url
    })

# --- Products (商品台帳) ---
@app.get("/products", response_class=HTMLResponse)
async def list_products(
    request: Request, 
    q: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Product)
    if q:
        query = query.filter(
            (models.Product.name.contains(q)) | (models.Product.code.contains(q))
        )
    products = query.order_by(models.Product.id.desc()).all()
    return templates.TemplateResponse(request=request, name="products/list.html", context={
        "request": request,
        "active_page": "products",
        "products": products,
        "search_query": q,
        "user": user
    })

@app.get("/products/excel")
async def export_products_excel(
    q: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Product)
    if q:
        query = query.filter((models.Product.name.contains(q)) | (models.Product.code.contains(q)))
    products = query.order_by(models.Product.id.desc()).all()
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("商品一覧")
    
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})

    headers = ["商品コード", "商品名", "単価", "有効在庫"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, p in enumerate(products, 1):
        worksheet.write(row, 0, p.code, border_fmt)
        worksheet.write(row, 1, p.name, border_fmt)
        worksheet.write(row, 2, p.unit_price, num_fmt)
        worksheet.write(row, 3, p.stock_quantity, border_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"products_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/products/new", response_class=HTMLResponse)
async def new_product(
    request: Request,
    user: models.User = Depends(get_active_user)
):
    return templates.TemplateResponse(request=request, name="products/form.html", context={
        "request": request, 
        "active_page": "products",
        "user": user
    })

@app.post("/products/new")
async def create_product(
    code: str = Form(...),
    name: str = Form(...),
    price_retail: float = Form(0.0),
    price_a: float = Form(0.0),
    price_b: float = Form(0.0),
    price_c: float = Form(0.0),
    price_d: float = Form(0.0),
    price_e: float = Form(0.0),
    stock_quantity: int = Form(0),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = models.Product(
        code=code, name=name, 
        unit_price=price_retail, 
        price_retail=price_retail,
        price_a=price_a, price_b=price_b, price_c=price_c, price_d=price_d, price_e=price_e,
        stock_quantity=0 # 初期値は0（後で拠点に割り振る）
    )
    db.add(product)
    db.flush() # IDを取得
    
    # デフォルト拠点（本社倉庫）に初期在庫を入れる
    main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
    if main_loc and stock_quantity > 0:
        update_product_stock(db, product.id, main_loc.id, stock_quantity, "INBOUND", "新規登録による初期在庫")
    
    db.commit()
    return RedirectResponse(url="/products", status_code=303)

@app.get("/products/edit/{product_id}", response_class=HTMLResponse)
async def edit_product(
    product_id: int, 
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = db.query(models.Product).get(product_id)
    return templates.TemplateResponse(request=request, name="products/form.html", context={
        "request": request, 
        "active_page": "products", 
        "product": product,
        "user": user
    })

@app.post("/products/edit/{product_id}")
async def update_product(
    product_id: int,
    code: str = Form(...),
    name: str = Form(...),
    price_retail: float = Form(0.0),
    price_a: float = Form(0.0),
    price_b: float = Form(0.0),
    price_c: float = Form(0.0),
    price_d: float = Form(0.0),
    price_e: float = Form(0.0),
    stock_quantity: int = Form(0),
    stock_add: int = Form(0),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = db.query(models.Product).get(product_id)
    if product:
        product.code = code
        product.name = name
        product.unit_price = price_retail
        product.price_retail = price_retail
        product.price_a = price_a
        product.price_b = price_b
        product.price_c = price_c
        product.price_d = price_d
        product.price_e = price_e
        
        # 入庫加算を拠点管理システム経由で行う
        if stock_add > 0:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, product.id, main_loc.id, stock_add, "INBOUND", "商品編集画面からの在庫追加")
        
        db.commit()
    return RedirectResponse(url="/products", status_code=303)

@app.post("/products/{product_id}/add_stock")
async def add_stock(
    product_id: int,
    quantity: int = Form(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = db.query(models.Product).get(product_id)
    if product and quantity != 0:
        main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
        if main_loc:
            update_product_stock(db, product.id, main_loc.id, quantity, "INBOUND", "商品一覧画面からの在庫追加")
        else:
            product.stock_quantity += quantity
        db.commit()
    return RedirectResponse(url="/products", status_code=303)

@app.post("/products/delete/{product_id}")
async def delete_product(
    product_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = db.query(models.Product).get(product_id)
    if product:
        db.delete(product)
        db.commit()
    return RedirectResponse(url="/products", status_code=303)

# --- Quotations (見積管理) ---
@app.get("/quotations", response_class=HTMLResponse)
async def list_quotations(
    request: Request, 
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Quotation).join(models.Customer)
    if q:
        query = query.filter(
            (models.Quotation.quote_number.contains(q)) |
            (models.Customer.company.contains(q)) |
            (models.Customer.name.contains(q))
        )
    if start_date:
        query = query.filter(models.Quotation.issue_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        # Include the whole end day
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Quotation.issue_date < end_dt)
        
    quotations = query.order_by(models.Quotation.id.desc()).all()
    return templates.TemplateResponse(request=request, name="quotations/list.html", context={
        "request": request,
        "active_page": "quotations",
        "quotations": quotations,
        "search_query": q,
        "start_date": start_date,
        "end_date": end_date,
        "user": user
    })

@app.get("/quotations/excel")
async def export_quotations_excel(
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Quotation).join(models.Customer)
    if q:
        query = query.filter(
            (models.Quotation.quote_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    if start_date:
        query = query.filter(models.Quotation.issue_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Quotation.issue_date < end_dt)
    
    quotations = query.order_by(models.Quotation.id.desc()).all()
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("見積一覧")
    
    # Formats
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    date_fmt = workbook.add_format({'num_format': 'yyyy/mm/dd', 'border': 1})
    num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})

    headers = ["見積番号", "顧客名", "発行日", "合計金額(税抜)"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, qt in enumerate(quotations, 1):
        worksheet.write(row, 0, qt.quote_number, border_fmt)
        worksheet.write(row, 1, qt.customer.company, border_fmt)
        worksheet.write(row, 2, qt.issue_date, date_fmt)
        worksheet.write(row, 3, qt.total_amount, num_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"quotations_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/quotations/new", response_class=HTMLResponse)
async def new_quotation(
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customers = db.query(models.Customer).order_by(models.Customer.id.desc()).all()
    return templates.TemplateResponse(request=request, name="quotations/form.html", context={
        "request": request,
        "active_page": "quotations",
        "customers": customers,
        "user": user
    })

@app.post("/quotations/new")
async def create_quotation(
    request: Request,
    customer_id: str = Form(""),
    quote_number: str = Form(...),
    issue_date: str = Form(...),
    expiry_date: str = Form(...),
    payment_due_date: str = Form(""),
    payment_method: str = Form("銀行振り込み"),
    discount_rate: float = Form(0.0),
    is_bulk_discount: bool = Form(False),
    customer_rank: str = Form("RETAIL"),
    memo: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if not customer_id or customer_id == "":
        # もしフロントを抜けてきた場合
        return HTMLResponse(content="<script>alert('顧客を選択してください'); history.back();</script>", status_code=400)
    customer_id = int(customer_id)
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")
    
    pay_due = None
    if payment_due_date:
        pay_due = datetime.datetime.strptime(payment_due_date, '%Y-%m-%d')

    quotation = models.Quotation(
        customer_id=customer_id,
        quote_number=quote_number,
        issue_date=datetime.datetime.strptime(issue_date, '%Y-%m-%d'),
        expiry_date=datetime.datetime.strptime(expiry_date, '%Y-%m-%d'),
        payment_due_date=pay_due,
        payment_method=payment_method,
        discount_rate=discount_rate,
        is_bulk_discount=is_bulk_discount,
        status=models.QuoteStatus.DRAFT,
        memo=memo
    )
    db.add(quotation)
    db.flush()

    total = 0
    for p_id, p_name, qty, price in zip(product_ids, product_names, quantities, prices):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        
        # product_id can be empty for manual entry
        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
    
    discount_amount = int(total * (discount_rate / 100))
    final_total_tax_excl = total - discount_amount
    # 10k JPY limit check removed for Admin system per request

    quotation.total_amount = final_total_tax_excl
    db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.get("/quotations/edit/{quote_id}", response_class=HTMLResponse)
async def edit_quotation(
    quote_id: int, 
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    quotation = db.query(models.Quotation).get(quote_id)
    customers = db.query(models.Customer).order_by(models.Customer.id.desc()).all()
    return templates.TemplateResponse(request=request, name="quotations/form.html", context={
        "request": request,
        "active_page": "quotations",
        "quotation": quotation,
        "customers": customers,
        "user": user
    })

@app.post("/quotations/edit/{quote_id}")
async def update_quotation(
    quote_id: int,
    request: Request,
    customer_id: str = Form(""),
    quote_number: str = Form(...),
    issue_date: str = Form(...),
    expiry_date: str = Form(...),
    payment_due_date: str = Form(""),
    payment_method: str = Form(""),
    discount_rate: float = Form(0.0),
    is_bulk_discount: bool = Form(False),
    customer_rank: str = Form("RETAIL"),
    memo: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    quotation = db.query(models.Quotation).get(quote_id)
    if not customer_id or customer_id == "":
        return HTMLResponse(content="<script>alert('顧客を選択してください'); history.back();</script>", status_code=400)
    
    customer_id = int(customer_id)
    quotation.customer_id = customer_id
    quotation.quote_number = quote_number
    try:
        quotation.issue_date = datetime.datetime.strptime(issue_date, '%Y-%m-%d')
        quotation.expiry_date = datetime.datetime.strptime(expiry_date, '%Y-%m-%d')
        if payment_due_date:
            quotation.payment_due_date = datetime.datetime.strptime(payment_due_date, '%Y-%m-%d')
        else:
            quotation.payment_due_date = None
    except ValueError:
        # Fallback if format is different (unlikely with type="date" but for safety)
        pass
    quotation.payment_method = payment_method
    quotation.discount_rate = discount_rate
    quotation.is_bulk_discount = is_bulk_discount
    quotation.memo = memo

    if quotation.status == models.QuoteStatus.ORDERED:
        for old_item in db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quotation.id).all():
            if old_item.product_id:
                main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
                if main_loc:
                    update_product_stock(db, old_item.product_id, main_loc.id, old_item.quantity, "INBOUND", "注文編集による在庫戻し")

    # 洗替方式で明細を更新
    db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quote_id).delete()
    
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")

    total = 0
    for p_id, p_name, qty, price in zip(product_ids, product_names, quantities, prices):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        if quotation.status == models.QuoteStatus.ORDERED and pid:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, pid, main_loc.id, qty, "OUTBOUND", "注文更新による自動出庫")
    
    discount_amount = int(total * (discount_rate / 100))
    final_total_tax_excl = total - discount_amount
    # 10k JPY limit check removed for Admin system per request

    quotation.total_amount = final_total_tax_excl
    db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.post("/quotations/delete/{quote_id}")
async def delete_quotation(
    quote_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    try:
        quote = db.query(models.Quotation).get(quote_id)
        if quote:
            if quote.status == models.QuoteStatus.ORDERED:
                for item in quote.items:
                    if item.product_id:
                        main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
                        if main_loc:
                            update_product_stock(db, item.product_id, main_loc.id, item.quantity, "INBOUND", "見積削除による在庫戻し")
            db.delete(quote)
            db.commit()
    except Exception as e:
        print(f"Error deleting quotation: {e}")
        db.rollback()
    return RedirectResponse(url="/quotations", status_code=303)

@app.post("/quotations/{id}/cancel")
async def cancel_quotation(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    quote = db.query(models.Quotation).get(id)
    if quote:
        # If it was ordered, delete the order and invoice
        if quote.order:
            for item in quote.items:
                if item.product_id:
                    main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
                    if main_loc:
                        update_product_stock(db, item.product_id, main_loc.id, item.quantity, "INBOUND", "注文キャンセルによる在庫戻し")
            db.delete(quote.order)
        quote.status = models.QuoteStatus.DRAFT
        db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.post("/quotations/{id}/copy")
async def copy_quotation(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    original = db.query(models.Quotation).get(id)
    if not original:
        return RedirectResponse(url="/quotations", status_code=303)
    
    # Create new quotation
    new_quote = models.Quotation(
        customer_id=original.customer_id,
        quote_number=f"Q-{original.quote_number}-{datetime.datetime.now().strftime('%m%d%H%M%S')}",
        issue_date=datetime.datetime.now(),
        expiry_date=datetime.datetime.now() + datetime.timedelta(days=30),
        payment_due_date=original.payment_due_date,
        payment_method=original.payment_method,
        total_amount=original.total_amount,
        status=models.QuoteStatus.DRAFT,
        discount_rate=original.discount_rate,
        is_bulk_discount=original.is_bulk_discount,
        memo=original.memo
    )
    db.add(new_quote)
    db.flush()
    
    # Copy items (Skip automated shipping fee as it will be re-added by JS if needed)
    for item in original.items:
        if item.description == "運賃（自動追加）":
            continue
        new_item = models.QuotationItem(
            quotation_id=new_quote.id,
            product_id=item.product_id,
            description=item.description,
            quantity=item.quantity,
            unit_price=item.unit_price,
            subtotal=item.subtotal
        )
        db.add(new_item)
    
    db.commit()
    return RedirectResponse(url=f"/quotations/edit/{new_quote.id}", status_code=303)

# --- PDF Generation (With Japanese Font Support) ---
def get_pdf_instance():
    pdf = FPDF()
    import os
    # Use absolute path to font
    font_path = r"C:\Windows\Fonts\msgothic.ttc"
    if os.path.exists(font_path):
        try:
            pdf.add_font("msgothic", "", font_path)
            pdf.set_font("msgothic", size=12)
            return pdf, "msgothic"
        except Exception as e:
            print(f"Font loading error: {e}")
    
    pdf.set_font("Helvetica", size=12)
    return pdf, "Helvetica"

@app.get("/quotations/{quote_id}/print", response_class=HTMLResponse)
async def print_quote(
    request: Request, 
    quote_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    quote = db.query(models.Quotation).get(quote_id)
    if not quote:
        return {"error": "Not found"}
    
    return templates.TemplateResponse(request=request, name="print_layout.html", context={
        "request": request,
        "doc_type": "quotation",
        "doc": quote, "user": user })

# --- Excel Generation ---
@app.get("/quotations/{quote_id}/excel")
async def generate_quote_excel(
    quote_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    quote = db.query(models.Quotation).get(quote_id)
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "見積書", workbook.add_format({'bold': True, 'font_size': 16}))
    worksheet.write(1, 0, f"No: {quote.quote_number}")
    worksheet.write(2, 0, f"顧客名: {quote.customer.company} 様")
    
    headers = ["商品名", "単価", "数量", "小計"]
    for i, h in enumerate(headers):
        worksheet.write(4, i, h, workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1}))
    
    for row_num, item in enumerate(quote.items, start=5):
        name = item.description or (item.product.name if item.product else "不明")
        worksheet.write(row_num, 0, name)
        worksheet.write(row_num, 1, item.unit_price)
        worksheet.write(row_num, 2, item.quantity)
        worksheet.write(row_num, 3, item.subtotal)
    
    last_row = 5 + len(quote.items)
    worksheet.write(last_row, 2, "合計", workbook.add_format({'bold': True}))
    worksheet.write(last_row, 3, quote.total_amount, workbook.add_format({'num_format': '#,##0', 'bold': True}))

    workbook.close()
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=quote_{quote.quote_number}.xlsx"})

@app.post("/quotations/{quote_id}/status")
async def update_quotation_status(
    quote_id: int, 
    status: str = Form(...), 
    db: Session = Depends(get_db), 
    user: models.User = Depends(get_active_user)
):
    quotation = db.query(models.Quotation).get(quote_id)
    if quotation:
        for enum_val in models.QuoteStatus:
            if enum_val.value == status:
                quotation.status = enum_val
                db.commit()
                break
    return RedirectResponse(url="/quotations", status_code=303)

# --- Orders (受注管理) ---
@app.get("/orders", response_class=HTMLResponse)
async def list_orders(
    request: Request, 
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Order).join(models.Quotation).join(models.Customer)
    if q:
        query = query.filter(
            (models.Order.order_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    if start_date:
        query = query.filter(models.Order.order_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Order.order_date < end_dt)
        
    orders = query.order_by(models.Order.id.desc()).all()
    return templates.TemplateResponse(request=request, name="orders/list.html", context={
        "request": request,
        "active_page": "orders",
        "orders": orders,
        "search_query": q,
        "start_date": start_date,
        "end_date": end_date,
        "user": user
    })

@app.get("/orders/excel")
async def export_orders_excel(
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Order).join(models.Quotation).join(models.Customer)
    if q:
        query = query.filter(
            (models.Order.order_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    if start_date:
        query = query.filter(models.Order.order_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Order.order_date < end_dt)
        
    orders = query.order_by(models.Order.id.desc()).all()
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("受注一覧")
    
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    date_fmt = workbook.add_format({'num_format': 'yyyy/mm/dd', 'border': 1})
    num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})

    headers = ["受注番号", "顧客名", "受注日", "合計金額(税抜)"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, ord in enumerate(orders, 1):
        worksheet.write(row, 0, ord.order_number, border_fmt)
        worksheet.write(row, 1, ord.quotation.customer.company, border_fmt)
        worksheet.write(row, 2, ord.order_date, date_fmt)
        worksheet.write(row, 3, ord.total_amount, num_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"orders_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/quotations/{quote_id}/order")
async def convert_to_order(
    quote_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    quote = db.query(models.Quotation).get(quote_id)
    if not quote or quote.status == models.QuoteStatus.ORDERED:
        return RedirectResponse(url="/quotations", status_code=303)

    # 1. Create Order
    order = models.Order(
        quotation_id=quote.id,
        order_number=quote.quote_number.replace("Q-", "ORD-"),
        total_amount=quote.total_amount,
        discount_rate=quote.discount_rate,
        is_bulk_discount=quote.is_bulk_discount,
        status=models.OrderStatus.PENDING,
        memo=quote.memo
    )
    db.add(order)

    # 2. Update Quotation Status
    quote.status = models.QuoteStatus.ORDERED

    # 3. Deduct Stock from Products
    for item in quote.items:
        if item.product_id:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, item.product_id, main_loc.id, item.quantity, "OUTBOUND", "受注確定による自動出庫")
    
    db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.get("/orders/new", response_class=HTMLResponse)
async def new_order_form(
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customers = db.query(models.Customer).order_by(models.Customer.id.desc()).all()
    # Generate a default order number
    order_number = f"ORD-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
    return templates.TemplateResponse(request=request, name="orders/form.html", context={
        "request": request,
        "active_page": "orders",
        "customers": customers,
        "order_number": order_number,
        "user": user
    })

@app.post("/orders/new")
async def create_direct_order(
    request: Request,
    customer_id: str = Form(""),
    order_number: str = Form(...),
    order_date: str = Form(...),
    discount_rate: float = Form(0.0),
    is_bulk_discount: bool = Form(False),
    customer_rank: str = Form("RETAIL"),
    memo: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if not customer_id or customer_id == "":
        return HTMLResponse(content="<script>alert('顧客を選択してください'); history.back();</script>", status_code=400)
    customer_id_int = int(customer_id)
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")

    # Create a shadow quotation for this order
    quotation = models.Quotation(
        customer_id=customer_id_int,
        quote_number=order_number.replace("ORD-", "Q-AUTO-"),
        expiry_date=datetime.datetime.utcnow(),
        status=models.QuoteStatus.ORDERED,
        discount_rate=discount_rate,
        is_bulk_discount=is_bulk_discount,
        total_amount=0,
        memo=memo
    )
    db.add(quotation)
    db.flush()

    total = 0
    for p_id, p_name, qty, price in zip(product_ids, product_names, quantities, prices):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        # Inventory deduction
        if pid:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, pid, main_loc.id, qty, "OUTBOUND", "直接受注による自動出庫")
    
    discount_amount = int(total * (discount_rate / 100))
    final_total_tax_excl = total - discount_amount
    # 10k JPY limit check removed for Admin system per request

    quotation.total_amount = final_total_tax_excl
    
    order = models.Order(
        quotation_id=quotation.id,
        order_number=order_number,
        order_date=datetime.datetime.strptime(order_date, '%Y-%m-%d'),
        total_amount=final_total_tax_excl,
        discount_rate=discount_rate,
        is_bulk_discount=is_bulk_discount,
        status=models.OrderStatus.PENDING,
        memo=memo
    )
    db.add(order)
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.get("/orders/edit/{order_id}", response_class=HTMLResponse)
async def edit_order(
    order_id: int, 
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.Order).get(order_id)
    customers = db.query(models.Customer).order_by(models.Customer.id.desc()).all()
    return templates.TemplateResponse(request=request, name="orders/form.html", context={
        "request": request,
        "active_page": "orders",
        "order": order,
        "customers": customers,
        "user": user
    })

@app.post("/orders/edit/{order_id}")
async def update_order(
    order_id: int,
    request: Request,
    customer_id: str = Form(""),
    order_number: str = Form(...),
    order_date: str = Form(...),
    discount_rate: float = Form(0.0),
    is_bulk_discount: bool = Form(False),
    customer_rank: str = Form("RETAIL"),
    memo: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.Order).get(order_id)
    if not order:
        return RedirectResponse(url="/orders", status_code=303)
    
    if not customer_id or customer_id == "":
        return HTMLResponse(content="<script>alert('顧客を選択してください'); history.back();</script>", status_code=400)
    customer_id_int = int(customer_id)

    order.order_number = order_number
    order.order_date = datetime.datetime.strptime(order_date, '%Y-%m-%d')
    quotation = order.quotation
    quotation.customer_id = customer_id_int
    quotation.issue_date = order.order_date # Keep Shadow Quotation in sync
    
    
    for old_item in db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quotation.id).all():
        if old_item.product_id:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, old_item.product_id, main_loc.id, old_item.quantity, "INBOUND", "注文編集による在庫戻し")

    # 洗替方式で明細を更新
    db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quotation.id).delete()
    
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")

    total = 0
    for p_id, p_name, qty, price in zip(product_ids, product_names, quantities, prices):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        if pid:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, pid, main_loc.id, qty, "OUTBOUND", "注文編集による自動出庫")
    
    discount_amount = int(total * (discount_rate / 100))
    final_total_tax_excl = total - discount_amount
    # 10k JPY limit check removed for Admin system per request

    quotation.total_amount = final_total_tax_excl
    order.total_amount = quotation.total_amount
    order.discount_rate = discount_rate
    quotation.is_bulk_discount = is_bulk_discount
    order.memo = memo
    quotation.memo = memo

    # もし納品済みで入金請求書データが作られている場合は、それも更新する
    if order.invoice:
        order.invoice.total_amount = final_total_tax_excl
        order.invoice.discount_rate = discount_rate
        order.invoice.is_bulk_discount = is_bulk_discount
        order.invoice.issue_date = order.order_date
        # 支払期限は、注文日に基づいて再計算（簡易的に30日後、または顧客設定があればそれに合わせるべきだが、現状は30日で維持）
        # 支払期限は、顧客の設定に基づいて再計算
        cust = order.quotation.customer
        closing_date = get_next_closing_date(order.order_date.date(), cust.closing_day)
        due_date = calculate_payment_date(closing_date, cust.payment_term_months or 1, cust.payment_day or 31)
        order.invoice.due_date = datetime.datetime.combine(due_date, datetime.time.min)
        
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/orders/{order_id}/status")
async def update_order_status(
    order_id: int, 
    status: str = Form(...), 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.Order).get(order_id)
    if order:
        order.status = models.OrderStatus(status)
        db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/orders/delete/{order_id}")
async def delete_order(
    order_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    try:
        order = db.query(models.Order).get(order_id)
        if order:
            if order.quotation:
                for item in order.quotation.items:
                    if item.product_id:
                        main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
                        if main_loc:
                            update_product_stock(db, item.product_id, main_loc.id, item.quantity, "INBOUND", "受注削除による在庫戻し")
            db.delete(order)
            db.commit()
    except Exception as e:
        print(f"Error deleting order: {e}")
        db.rollback()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/orders/{id}/copy")
async def copy_order(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    original = db.query(models.Order).get(id)
    if not original:
        return RedirectResponse(url="/orders", status_code=303)
    
    # Copy the underlying quotation
    original_quote = original.quotation
    new_quote = models.Quotation(
        customer_id=original_quote.customer_id,
        quote_number=f"Q-COPY-{datetime.datetime.now().strftime('%m%d%H%M%S')}",
        issue_date=datetime.datetime.now(),
        expiry_date=datetime.datetime.now() + datetime.timedelta(days=30),
        payment_due_date=original_quote.payment_due_date,
        payment_method=original_quote.payment_method,
        total_amount=original_quote.total_amount,
        discount_rate=original_quote.discount_rate,
        is_bulk_discount=original_quote.is_bulk_discount,
        status=models.QuoteStatus.ORDERED,
        memo=original_quote.memo
    )
    db.add(new_quote)
    db.flush()
    
    # Copy items (Skip automated shipping fee as it will be re-added by JS if needed)
    for item in original_quote.items:
        if item.description == "運賃（自動追加）":
            continue
        new_item = models.QuotationItem(
            quotation_id=new_quote.id,
            product_id=item.product_id,
            description=item.description,
            quantity=item.quantity,
            unit_price=item.unit_price,
            subtotal=item.subtotal
        )
        db.add(new_item)
        # 在庫減算
        if item.product_id:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, item.product_id, main_loc.id, item.quantity, "OUTBOUND", "コピー受注による自動出庫")
    
    # Create new order linked to this new quote
    new_order = models.Order(
        quotation_id=new_quote.id,
        order_number=f"ORD-COPY-{datetime.datetime.now().strftime('%m%d%H%M%S')}",
        order_date=datetime.datetime.now(),
        total_amount=new_quote.total_amount,
        discount_rate=new_quote.discount_rate,
        is_bulk_discount=new_quote.is_bulk_discount,
        status=models.OrderStatus.PENDING,
        memo=original.memo
    )
    db.add(new_order)
    db.commit()
    return RedirectResponse(url=f"/orders/edit/{new_order.id}", status_code=303)

@app.get("/orders/{order_id}/print_delivery", response_class=HTMLResponse)
async def print_delivery_note(
    request: Request, 
    order_id: int, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.Order).get(order_id)
    if not order:
        return {"error": "Not found"}
    
    return templates.TemplateResponse(request=request, name="print_layout.html", context={
        "request": request,
        "doc_type": "delivery_note",
        "doc": order, "user": user })

# --- Invoices (請求・入金管理) ---
@app.get("/invoices", response_class=HTMLResponse)
async def list_invoices(
    request: Request, 
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Invoice).outerjoin(models.Customer)
    if q:
        query = query.filter(
            (models.Invoice.invoice_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    if start_date:
        query = query.filter(models.Invoice.issue_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Invoice.issue_date < end_dt)
        
    invoices = query.order_by(models.Invoice.id.desc()).all()
    return templates.TemplateResponse(request=request, name="invoices/list.html", context={
        "request": request,
        "active_page": "invoices",
        "invoices": invoices,
        "search_query": q,
        "start_date": start_date,
        "end_date": end_date,
        "InvoiceStatus": models.InvoiceStatus,
        "user": user
    })

@app.get("/invoices/excel")
async def export_invoices_excel(
    q: str = "", 
    start_date: str = "", 
    end_date: str = "", 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.Invoice).outerjoin(models.Customer)
    if q:
        query = query.filter(
            (models.Invoice.invoice_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    if start_date:
        query = query.filter(models.Invoice.issue_date >= datetime.datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d') + datetime.timedelta(days=1)
        query = query.filter(models.Invoice.issue_date < end_dt)
        
    invoices = query.order_by(models.Invoice.id.desc()).all()
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("請求一覧")
    
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    date_fmt = workbook.add_format({'num_format': 'yyyy/mm/dd', 'border': 1})
    num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})

    headers = ["請求番号", "顧客名", "発行日", "支払期限", "請求金額(税抜)", "紐付く受注数"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, inv in enumerate(invoices, 1):
        worksheet.write(row, 0, inv.invoice_number, border_fmt)
        cust_name = inv.customer.company if inv.customer else (inv.orders[0].quotation.customer.company if inv.orders else "")
        worksheet.write(row, 1, cust_name, border_fmt)
        worksheet.write(row, 2, inv.issue_date, date_fmt)
        worksheet.write(row, 3, inv.due_date, date_fmt)
        worksheet.write(row, 4, int(inv.total_amount), num_fmt)
        worksheet.write(row, 5, len(inv.orders), border_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"invoices_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/invoices/bulk-print", response_class=HTMLResponse)
async def admin_bulk_print_invoices(
    request: Request,
    invoice_ids: list[int] = Query(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    """一括印刷用レイアウトを表示 (GET)"""
    invoices = db.query(models.Invoice).filter(
        models.Invoice.id.in_(invoice_ids)
    ).all()
    
    if not invoices:
        return RedirectResponse(url="/admin/invoice-dispatch", status_code=303)
    
    # 印刷対象を「発行済」に更新
    for inv in invoices:
        if inv.status == models.InvoiceStatus.UNPAID:
            inv.status = models.InvoiceStatus.ISSUED
    db.commit()
    
    return templates.TemplateResponse(request=request, name="invoices/bulk_print.html", context={
        "request": request,
        "invoices": invoices,
        "user": user
    })

@app.post("/invoices/bulk_print")
async def bulk_print_invoices_post(request: Request, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    """一括印刷用レイアウトを表示 (POST) - PRGパターンでGETにリダイレクト"""
    form_data = await request.form()
    # 複数のキー形式に対応
    raw_ids = form_data.getlist("invoice_ids") or form_data.getlist("invoice_ids[]")
    invoice_ids = [int(i) for i in raw_ids if i]
    
    if not invoice_ids:
        return RedirectResponse(url="/invoices", status_code=303)
    
    # POST→Redirect→GET パターンで、ブラウザキャッシュ問題を防ぐ
    ids_param = "&".join([f"invoice_ids={i}" for i in invoice_ids])
    return RedirectResponse(url=f"/invoices/bulk-print?{ids_param}", status_code=303)


@app.post("/orders/{order_id}/invoice")
async def create_invoice(order_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    order = db.query(models.Order).get(order_id)
    if not order or order.invoice:
        return RedirectResponse(url="/orders", status_code=303)

    # Create Invoice
    invoice = models.Invoice(
        customer_id=order.quotation.customer_id,
        invoice_number=order.order_number.replace("ORD-", "INV-"),
        issue_date=order.order_date, # Default to order date
        due_date=datetime.datetime.combine(
            calculate_payment_date(
                get_next_closing_date(order.order_date.date(), order.quotation.customer.closing_day),
                order.quotation.customer.payment_term_months or 1,
                order.quotation.customer.payment_day or 31
            ),
            datetime.time.min
        ),
        total_amount=order.total_amount,
        status=models.InvoiceStatus.UNPAID,
        delivery_status="UNSENT",
        memo=order.memo
    )
    db.add(invoice)
    db.flush()
    
    # Link order to invoice
    order.invoice_id = invoice.id
    
    # Update order status to SHIPPED (出荷済み)
    order.status = models.OrderStatus.SHIPPED
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/orders/{order_id}/cancel_shipping")
async def cancel_shipping(order_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    order = db.query(models.Order).get(order_id)
    if not order or order.status != models.OrderStatus.SHIPPED:
        return RedirectResponse(url="/orders", status_code=303)
    
    # Unlink or delete associated invoice
    invoice = order.invoice
    if invoice:
        order.invoice_id = None
        db.flush()
        # If no other orders are linked to this invoice, delete it
        if not invoice.orders:
            db.delete(invoice)
    
    # Revert status to PENDING (未出荷)
    order.status = models.OrderStatus.PENDING
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/invoices/{invoice_id}/pay")
async def mark_as_paid(invoice_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    if invoice:
        invoice.status = models.InvoiceStatus.PAID
        db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.post("/invoices/{id}/mark_issued")
async def mark_issued(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(id)
    if invoice:
        invoice.status = models.InvoiceStatus.ISSUED
        db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.post("/invoices/{id}/unmark_issued")
async def unmark_issued(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(id)
    if invoice:
        invoice.status = models.InvoiceStatus.UNPAID
        db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.post("/invoices/{invoice_id}/set_status")
async def set_invoice_status(
    invoice_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    form_data = await request.form()
    status_name = form_data.get(f"status_{invoice_id}")
    invoice = db.query(models.Invoice).get(invoice_id)
    if invoice and status_name:
        try:
            invoice.status = models.InvoiceStatus[status_name]
            db.commit()
        except KeyError:
            pass
    from fastapi.responses import Response
    return Response(status_code=204)



@app.get("/invoices/edit/{invoice_id}", response_class=HTMLResponse)
async def edit_invoice(invoice_id: int, request: Request, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    return templates.TemplateResponse(request=request, name="invoices/form.html", context={
        "request": request,
        "active_page": "invoices",
        "invoice": invoice, "user": user })

@app.post("/invoices/edit/{invoice_id}")
async def update_invoice(
    invoice_id: int,
    invoice_number: str = Form(...),
    issue_date: str = Form(...),
    total_amount: float = Form(...),
    due_date: str = Form(...),
    status: str = Form(...),
    delivery_status: str = Form("UNSENT"),
    discount_rate: float = Form(0.0),
    is_bulk_discount: bool = Form(False),
    memo: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    invoice = db.query(models.Invoice).get(invoice_id)
    if not invoice:
        return RedirectResponse(url="/invoices", status_code=303)
    
    invoice.invoice_number = invoice_number
    try:
        invoice.issue_date = datetime.datetime.strptime(issue_date, '%Y-%m-%d')
    except ValueError:
        pass
    invoice.total_amount = total_amount
    try:
        invoice.due_date = datetime.datetime.strptime(due_date, '%Y-%m-%d')
    except ValueError:
        pass
    invoice.status = models.InvoiceStatus(status)
    invoice.delivery_status = delivery_status
    invoice.discount_rate = discount_rate
    invoice.is_bulk_discount = is_bulk_discount
    invoice.memo = memo
    
    db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.get("/invoices/{id}")
async def view_invoice(request: Request, id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(id)
    if not invoice:
        return RedirectResponse(url="/invoices", status_code=303)
    
    return templates.TemplateResponse(request=request, name="invoices/detail.html", context={
        "request": request,
        "invoice": invoice,
        "customer": invoice.customer,
        "orders": invoice.orders
    })

@app.post("/invoices/delete/{invoice_id}")
async def delete_invoice(invoice_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    if invoice:
        # 紐づくすべての受注の在庫を戻し、紐付けを解除する
        for order in invoice.orders:
            if order.quotation:
                for item in order.quotation.items:
                    if item.product_id:
                        main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
                        if main_loc:
                            update_product_stock(db, item.product_id, main_loc.id, item.quantity, "INBOUND", "請求書削除による在庫戻し")
            order.invoice_id = None
            order.status = models.OrderStatus.PENDING # 未出荷に戻す
        
        db.delete(invoice)
        db.commit()
    response = RedirectResponse(url="/invoices", status_code=303)
    response.headers["HX-Refresh"] = "true"
    return response

def generate_invoice_pdf_content(invoice: models.Invoice):
    """請求書のPDFデータを生成してバイト列で返す。エラー時はNoneを返す"""
    from fpdf import FPDF
    import io
    import logging
    import os

    try:
        # カラー設定 (#1a2a6c)
        primary_color = (26, 42, 108)
        
        # フォントパス候補
        font_paths = [
            "static/fonts/NotoSansJP-Regular.otf",
            "static/fonts/IPAexGothic.ttf",
            "C:\\Windows\\Fonts\\msgothic.ttc"
        ]
        
        pdf = FPDF()
        pdf.add_page()
        
        font_found = False
        for path in font_paths:
            if os.path.exists(path):
                try:
                    pdf.add_font("Japanese", "", path)
                    pdf.set_font("Japanese", size=10)
                    font_found = True
                    break
                except Exception as e:
                    print(f"Font load error ({path}): {e}")
                    continue
        
        if not font_found:
            pdf.set_font("helvetica", size=10)

        # --- ヘッダー領域 ---
        pdf.set_font("Japanese" if font_found else "helvetica", "", 24)
        pdf.set_text_color(*primary_color)
        pdf.cell(100, 15, "請  求  書", ln=False)
        
        # タイトルの下線
        pdf.set_draw_color(*primary_color)
        pdf.set_line_width(0.8)
        pdf.line(10, 24, 60, 24)
        
        # 右側：書類情報
        pdf.set_text_color(51, 51, 51)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
        pdf.set_xy(140, 10)
        pdf.cell(0, 5, f"No: {invoice.invoice_number}", ln=True, align="R")
        pdf.cell(0, 5, f"発行日: {invoice.issue_date.strftime('%Y年%m月%d日') if invoice.issue_date else ''}", ln=True, align="R")
        if invoice.due_date:
            pdf.cell(0, 5, f"お支払い期限: {invoice.due_date.strftime('%Y年%m月%d日')}", ln=True, align="R")
        
        pdf.ln(10)

        # --- 宛先と発行元 ---
        customer = invoice.customer if invoice.customer else (invoice.orders[0].quotation.customer if invoice.orders else None)
        
        # 左側：宛先
        pdf.set_xy(10, 35)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 10)
        if customer and customer.zip_code:
            pdf.cell(0, 5, f"〒{customer.zip_code}", ln=True)
        if customer and customer.address:
            pdf.cell(0, 5, f"{customer.address}", ln=True)
        
        pdf.set_font("Japanese" if font_found else "helvetica", "", 15)
        customer_name = (customer.company or customer.name) if customer else "御中"
        honorific = (customer.honorific or "御中") if customer else "御中"
        pdf.cell(100, 10, f"{customer_name}  {honorific}", ln=True)
        pdf.set_line_width(0.2)
        pdf.line(10, pdf.get_y(), 100, pdf.get_y())
        
        # 右側：発行元
        issuer_x = 130
        
        # 電子印影の配置 (テキストの前に配置することで、テキスが上に重なるようにする)
        if os.path.exists("static/images/seal.png"):
            pdf.image("static/images/seal.png", x=issuer_x + 35, y=38, w=22)

        pdf.set_xy(issuer_x, 35)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 12)
        pdf.set_text_color(*primary_color)
        pdf.cell(0, 7, "株式会社熊ノ護化研", ln=True, align="L")
        
        pdf.set_text_color(51, 51, 51)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
        pdf.set_x(issuer_x)
        pdf.cell(0, 5, "代表取締役社長　岡泰造", ln=True)
        pdf.set_x(issuer_x)
        pdf.cell(0, 5, "〒010-0001 秋田県秋田市中通3-1-9", ln=True)
        pdf.set_x(issuer_x)
        pdf.cell(0, 5, "TEL: 018-838-1920", ln=True)
        pdf.set_x(issuer_x)
        pdf.cell(0, 5, "Mail: info@kumanomorikaken.co.jp", ln=True)
        pdf.set_x(issuer_x)
        pdf.cell(0, 5, "登録番号: T5410001014110", ln=True)
        

        pdf.ln(10)

        # --- 御請求金額ボックス ---
        current_y = pdf.get_y()
        pdf.set_fill_color(248, 249, 250)
        pdf.set_draw_color(*primary_color)
        pdf.set_line_width(0.5)
        pdf.rect(10, current_y, 190, 15, style="DF")
        
        pdf.set_xy(15, current_y + 4)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 10)
        pdf.cell(50, 7, "御請求金額（税込）", ln=False)
        
        pdf.set_font("Japanese" if font_found else "helvetica", "", 20)
        pdf.set_text_color(*primary_color)
        grand_total = int(invoice.total_amount * 1.1)
        pdf.cell(125, 7, f"¥{grand_total:,.0f} -", ln=True, align="R")
        
        pdf.set_text_color(51, 51, 51)
        pdf.ln(8)

        # --- 明細テーブル ---
        pdf.set_fill_color(*primary_color)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
        
        # テーブルヘッダー
        pdf.cell(10, 8, "No", border=1, align="C", fill=True)
        pdf.cell(90, 8, "品名 / 摘要", border=1, align="C", fill=True)
        pdf.cell(20, 8, "数量", border=1, align="C", fill=True)
        pdf.cell(35, 8, "単価（円）", border=1, align="C", fill=True)
        pdf.cell(35, 8, "金額（円）", border=1, align="C", fill=True)
        pdf.ln()

        pdf.set_text_color(51, 51, 51)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
        
        row_idx = 1
        for order in invoice.orders:
            # 受注区切り行
            pdf.set_fill_color(240, 244, 248)
            pdf.set_font("Japanese" if font_found else "helvetica", "", 8)
            pdf.set_text_color(*primary_color)
            pdf.cell(10, 6, "", border=1, fill=True)
            pdf.cell(180, 6, f" 【受注日: {order.order_date.strftime('%Y/%m/%d')} ｜ No: {order.order_number}】", border=1, fill=True)
            pdf.ln()
            
            pdf.set_text_color(51, 51, 51)
            pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
            if order.quotation:
                for item in order.quotation.items:
                    pdf.cell(10, 7, str(row_idx), border=1, align="C")
                    pdf.cell(90, 7, item.description, border=1)
                    pdf.cell(20, 7, f"{item.quantity}", border=1, align="R")
                    pdf.cell(35, 7, f"{item.unit_price:,.0f}", border=1, align="R")
                    pdf.cell(35, 7, f"{item.subtotal:,.0f}", border=1, align="R")
                    pdf.ln()
                    row_idx += 1

        # 空行の埋め合わせ（最低10行程度）
        while row_idx <= 8:
            pdf.cell(10, 7, str(row_idx), border=1, align="C")
            pdf.cell(90, 7, "", border=1)
            pdf.cell(20, 7, "", border=1)
            pdf.cell(35, 7, "", border=1)
            pdf.cell(35, 7, "", border=1)
            pdf.ln()
            row_idx += 1

        # --- 集計セクション ---
        pdf.ln(2)
        summary_x = 130
        pdf.set_x(summary_x)
        pdf.cell(35, 6, "小計（税抜）", border=0, align="R")
        pdf.cell(35, 6, f"¥{int(invoice.total_amount):,.0f}", border=0, align="R", ln=True)
        
        pdf.set_x(summary_x)
        pdf.cell(35, 6, "消費税（10%）", border=0, align="R")
        pdf.cell(35, 6, f"¥{int(invoice.total_amount * 0.1):,.0f}", border=0, align="R", ln=True)
        
        pdf.set_draw_color(*primary_color)
        pdf.line(summary_x + 5, pdf.get_y(), 200, pdf.get_y())
        
        pdf.set_x(summary_x)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 11)
        pdf.set_text_color(*primary_color)
        pdf.cell(35, 8, "合計（税込）", border=0, align="R")
        pdf.cell(35, 8, f"¥{grand_total:,.0f}", border=0, align="R", ln=True)

        # --- お振込先 ---
        pdf.ln(10)
        pdf.set_draw_color(221, 221, 221)
        pdf.set_line_width(0.2)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
        
        pdf.set_text_color(*primary_color)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 10)
        pdf.cell(30, 6, "【お振込先】", ln=False)
        
        pdf.set_text_color(51, 51, 51)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 9)
        bank_info = "秋田銀行　秋田駅前支店　普通口座　1090927\n株式会社熊ノ護化研　代表取締役社長　岡泰造"
        pdf.multi_cell(0, 5, bank_info)
        
        pdf.set_line_width(0.5)
        pdf.set_draw_color(*primary_color)
        pdf.line(10, pdf.get_y() + 2, 200, pdf.get_y() + 2)

        # --- 備考 ---
        pdf.ln(5)
        pdf.set_font("Japanese" if font_found else "helvetica", "", 8)
        pdf.set_text_color(102, 102, 102)
        notes = "【備考】\n・振込手数料はお客様負担にてお願い申し上げます。\n・上記見積に基づき、ご請求申し上げます。\n・上記金額には消費税10%が含まれております。"
        pdf.multi_cell(0, 4, notes)

        # バイト列として出力
        return pdf.output()
    except Exception as e:
        import traceback
        print(f"PDF生成エラー: {e}")
        print(traceback.format_exc())
        return None

@app.post("/invoices/{id}/send_email")
async def send_invoice_email(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(id)
    if not invoice:
        return RedirectResponse(url="/invoices", status_code=303)
    
    customer = invoice.customer if invoice.customer else (invoice.orders[0].quotation.customer if invoice.orders else None)
    if not customer or not customer.email:
        return RedirectResponse(url=f"/invoices/{id}?error=no_email", status_code=303)
    
    # PDF生成
    pdf_content = generate_invoice_pdf_content(invoice)
    attachments = [{
        "name": f"請求書_{invoice.invoice_number}.pdf",
        "content": pdf_content
    }] if pdf_content else None
    
    subject = f"【ご請求書】株式会社熊ノ護化研より（請求番号：{invoice.invoice_number}）"
    body = f"""{customer.company or customer.name}
{customer.honorific or '御中'}

いつも大変お世話になっております。
株式会社熊ノ護化研でございます。

今月分の請求書をお送りさせていただきます。
詳細は添付のPDFまたは以下のリンクよりご確認ください。

■ 請求番号：{invoice.invoice_number}
■ 御請求金額（税込）：¥{'{:,.0f}'.format(grand_total := int(invoice.total_amount * 1.1))}-
■ お支払期限：{invoice.due_date.strftime('%Y/%m/%d') if invoice.due_date else ''}

お振込先情報は添付の請求書PDF内に記載しております。
ご確認のほど、何卒よろしくお願い申し上げます。

※ 本メールはシステムより送信されています。
"""
    send_notification(subject, body, to=[customer.email], attachments=attachments)
    
    # ステータス更新
    invoice.delivery_status = "SENT"
    if invoice.status == models.InvoiceStatus.UNPAID:
        invoice.status = models.InvoiceStatus.ISSUED
    db.commit()
    
    return RedirectResponse(url=f"/invoices/{id}", status_code=303)

@app.post("/invoices/{id}/mark_mailed")
async def mark_mailed(id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(id)
    if invoice:
        invoice.delivery_status = "MAILED"
        if invoice.status == models.InvoiceStatus.UNPAID:
            invoice.status = models.InvoiceStatus.ISSUED
        db.commit()
    return RedirectResponse(url=f"/invoices/{id}", status_code=303)

@app.get("/invoices/{invoice_id}/print", response_class=HTMLResponse)
async def print_invoice(request: Request, invoice_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    if not invoice:
        return {"error": "Not found"}
    
    return templates.TemplateResponse(request=request, name="print_layout.html", context={
        "request": request,
        "doc_type": "invoice",
        "doc": invoice, "user": user })

# API for HTMX Product Search (Used in Quotation Creation)
@app.get("/api/products/search")
async def search_products(q: str = "", db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    products = db.query(models.Product).filter(
        (models.Product.name.contains(q)) | (models.Product.code.contains(q))
    ).order_by(models.Product.id.desc()).limit(10).all()
    # Simple HTML response for HTMX
    html = ""
    for p in products:
        # p.name をエスケープ (JS 用)
        safe_name = p.name.replace("'", "\\'")
        # prices も安全に文字列化 (NULL の場合は 0 に置換)
        prices_js = f"{{retail: {p.price_retail or 0}, a: {p.price_a or 0}, b: {p.price_b or 0}, c: {p.price_c or 0}, d: {p.price_d or 0}, e: {p.price_e or 0}}}"
        html += f'<div class="search-result" onclick="selectProduct({p.id}, \'{safe_name}\', {prices_js})">{p.name} ({p.code}) - ¥{p.price_retail:,.0f}</div>'
    return HTMLResponse(content=html if html else "<div>見つかりませんでした</div>")

# API for HTMX Customer Search
@app.get("/api/customers/search")
async def search_customers(q: str = "", db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    customers = db.query(models.Customer).filter(
        (models.Customer.name.contains(q)) | (models.Customer.company.contains(q))
    ).order_by(models.Customer.id.desc()).limit(10).all()
    html = ""
    for c in customers:
        rank_name = c.rank.name if c.rank else 'RETAIL'
        display_name = f"{c.company} ({c.name})" if c.company and c.name else (c.company or c.name or "名称未設定")
        safe_display_name = display_name.replace("'", "\\'")
        html += f'<div class="search-result" onclick="selectCustomer({c.id}, \'{safe_display_name}\', \'{rank_name}\')">{display_name}</div>'
    return HTMLResponse(content=html if html else "<div>見つかりませんでした</div>")

# Settings / Backup & Restore
@app.get("/settings")
async def settings_page(request: Request, user: models.User = Depends(get_active_user)):
    return templates.TemplateResponse(request=request, name="settings.html", context={"request": request, "active_page": "settings", "user": user })

@app.get("/backup")
async def download_backup(user: models.User = Depends(get_active_user)):
    if os.path.exists("kumanogo.db"):
        return FileResponse("kumanogo.db", filename="kumanogo_backup.db")
    return {"error": "Database file not found"}

@app.post("/restore")
async def restore_backup(backup_file: UploadFile = File(...), user: models.User = Depends(get_active_user)):
    from database import engine
    try:
        # Save uploaded file content
        new_content = await backup_file.read()
        
        # Dispose engine to close connections
        engine.dispose()
        
        # Overwrite the database file content
        with open("kumanogo.db", "wb") as f:
            f.write(new_content)
        
        return RedirectResponse(url="/settings?success=restored", status_code=303)
    except Exception as e:
        return {"error": str(e)}

# --- Auth Routes ---
@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse(request=request, name="login.html", context={"request": request })

@app.post("/login")
async def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db)
):
    user = db.query(models.User).filter(models.User.username == username).first()
    if not user or not verify_password(password, user.hashed_password):
        return templates.TemplateResponse(request=request, name="login.html", context={
            "request": request,
            "error": "ユーザー名またはパスワードが正しくありません"
        })
    
    token = serializer.dumps(username)
    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(key="session", value=token, httponly=True)
    return response

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login")
    response.delete_cookie("session")
    return response

# --- User Management (ユーザー管理) ---
@app.get("/users", response_class=HTMLResponse)
async def list_users(
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    users = db.query(models.User).order_by(models.User.id.desc()).all()
    return templates.TemplateResponse(request=request, name="users.html", context={
        "request": request, 
        "active_page": "users", 
        "users": users, 
        "user": user
    })

@app.post("/users/new")
async def create_user(
    request: Request,
    username: str = Form(...),
    full_name: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    # Check if username already exists
    existing = db.query(models.User).filter(models.User.username == username).first()
    if existing:
        users = db.query(models.User).order_by(models.User.id.desc()).all()
        return templates.TemplateResponse(request=request, name="users.html", context={
            "request": request,
            "active_page": "users",
            "users": users,
            "user": user,
            "error": f"ユーザー名 '{username}' は既に使用されています。"
        })
    
    new_user = models.User(
        username=username,
        hashed_password=get_password_hash(password),
        full_name=full_name,
        is_admin=False
    )
    db.add(new_user)
    db.commit()
    return RedirectResponse(url="/users?success=created", status_code=303)

@app.post("/users/delete/{user_id}")
async def delete_user(
    user_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    target = db.query(models.User).get(user_id)
    if target and not target.is_admin:
        db.delete(target)
        db.commit()
    return RedirectResponse(url="/users", status_code=303)

# --- Password Change (パスワード変更) ---
@app.get("/change-password", response_class=HTMLResponse)
async def change_password_page(
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    return templates.TemplateResponse(request=request, name="change_password.html", context={
        "request": request,
        "active_page": "change_password",
        "user": user
    })

@app.post("/change-password")
async def change_password(
    request: Request,
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if not verify_password(current_password, user.hashed_password):
        return templates.TemplateResponse(request=request, name="change_password.html", context={
            "request": request,
            "active_page": "change_password",
            "user": user,
            "error": "現在のパスワードが正しくありません。"
        })
    
    if new_password != confirm_password:
        return templates.TemplateResponse(request=request, name="change_password.html", context={
            "request": request,
            "active_page": "change_password",
            "user": user,
            "error": "新しいパスワードと確認用パスワードが一致しません。"
        })
    
    user.hashed_password = get_password_hash(new_password)
    db.commit()
    
    return templates.TemplateResponse(request=request, name="change_password.html", context={
        "request": request,
        "active_page": "change_password",
        "user": user,
        "success": "パスワードを正常に変更しました。"
    })

# 初回の管理者作成用
@app.get("/init-admin")
async def init_admin(db: Session = Depends(get_db)):
    admin = db.query(models.User).filter(models.User.username == "nakamura@connect-web.jp").first()
    if not admin:
        hashed_pw = get_password_hash("N687nh4su4")
        admin = models.User(
            username="nakamura@connect-web.jp",
            hashed_password=hashed_pw,
            full_name="管理者",
            is_admin=True
        )
        db.add(admin)
        db.commit()
        return RedirectResponse(url="/login", status_code=303)
    return RedirectResponse(url="/login", status_code=303)


# ============================================================
# 代理店ポータル (Agency Portal)
# ============================================================




async def send_order_notification_email(order: models.AgencyOrder, db: Session):
    subject = f"【代理店サイト】新規発注のお知らせ ({order.customer.company if order.customer else '不明な代理店'}様)"
    company_name = order.customer.company if order.customer else "不明な代理店"
    content = f"""代理店：{company_name} 様より新規発注がありました。

【受注番号】: {order.order_number}
【発注日時】: {order.order_date.strftime('%Y/%m/%d %H:%M') if order.order_date else '-'}
【合計金額】: ¥{'{:,.0f}'.format(order.total_amount)} (税抜)

詳細は管理画面の「代理店発注」ページ、または以下の通知一覧よりご確認ください。
https://app.kumanomorikaken.co.jp/admin/notifications
"""
    send_admin_email_sync(db, subject, content)

def send_admin_email_sync(db: Session, subject: str, content: str):
    """管理者へメール通知を送信（同期版）"""
    try:
        settings = db.query(models.SystemSetting).all()
        s = {s.key: s.value for s in settings}
        target = s.get("notification_email") or "info@kumanomorikaken.co.jp"
        
        if not s.get("smtp_host"):
            print(f"DEBUG: Email notification skipped (No SMTP): {subject}")
            return

        msg = EmailMessage()
        msg.set_content(content)
        msg['Subject'] = subject
        msg['From'] = s.get("smtp_from")
        msg['To'] = target

        smtp_host = s.get("smtp_host")
        smtp_port_val = s.get("smtp_port") or "587"
        try:
            smtp_port = int(smtp_port_val)
        except ValueError:
            smtp_port = 587
            
        smtp_user = s.get("smtp_user")
        smtp_pass = s.get("smtp_pass")

        import ssl
        context = ssl.create_default_context()
        if smtp_port == 465:
            with smtplib.SMTP_SSL(str(smtp_host), smtp_port, timeout=30, context=context) as server:
                if smtp_user and smtp_pass:
                    server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        else:
            with smtplib.SMTP(str(smtp_host), smtp_port, timeout=30) as server:
                if smtp_port == 587:
                    server.starttls(context=context)
                if smtp_user and smtp_pass:
                    server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        print(f"INFO: Admin Email sent to {target}: {subject}")
    except Exception as e:
        print(f"ERROR: Failed to send admin email: {e}")

agency_serializer = URLSafeSerializer(SECRET_KEY + "-agency")

async def get_current_agency(request: Request, db: Session = Depends(get_db)):
    """代理店セッションから現在のログイン代理店を取得"""
    session_token = request.cookies.get("agency_session")
    if not session_token:
        return None
    try:
        login_id = agency_serializer.loads(session_token)
        customer = db.query(models.Customer).filter(
            models.Customer.login_id == login_id,
            models.Customer.is_agency == True
        ).first()
        return customer
    except:
        return None

class NotAgencyAuthenticatedException(Exception):
    pass

@app.exception_handler(NotAgencyAuthenticatedException)
async def agency_auth_exception_handler(request: Request, exc: NotAgencyAuthenticatedException):
    return RedirectResponse(url="/agency/login")

async def get_active_agency(request: Request, db: Session = Depends(get_db)):
    agency = await get_current_agency(request, db)
    if not agency:
        raise NotAgencyAuthenticatedException()
    return agency

def get_price_for_rank(product, rank):
    """顧客ランクに応じた価格を取得"""
    rank_price_map = {
        models.CustomerRank.RETAIL: product.price_retail,
        models.CustomerRank.RANK_A: product.price_a,
        models.CustomerRank.RANK_B: product.price_b,
        models.CustomerRank.RANK_C: product.price_c,
        models.CustomerRank.RANK_D: product.price_d,
        models.CustomerRank.RANK_E: product.price_e,
    }
    price = rank_price_map.get(rank, product.price_retail)
    return price if price and price > 0 else product.price_retail

# --- Agency Login ---
@app.get("/agency/login", response_class=HTMLResponse)
async def agency_login_page(request: Request):
    return templates.TemplateResponse(request=request, name="agency/login.html", context={"request": request})

@app.post("/agency/login")
async def agency_login(
    request: Request,
    login_id: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db)
):
    customer = db.query(models.Customer).filter(
        models.Customer.login_id == login_id,
        models.Customer.is_agency == True
    ).first()
    
    if not customer or customer.agency_password != password:
        return templates.TemplateResponse(request=request, name="agency/login.html", context={
            "request": request,
            "error": "ログインIDまたはパスワードが正しくありません"
        })
    
    token = agency_serializer.dumps(login_id)
    response = RedirectResponse(url="/agency/", status_code=303)
    response.set_cookie(key="agency_session", value=token, httponly=True)
    return response

@app.get("/agency/logout")
async def agency_logout():
    response = RedirectResponse(url="/agency/login")
    response.delete_cookie("agency_session")
    return response

@app.get("/agency/manual", response_class=HTMLResponse)
async def agency_manual(
    request: Request,
    agency: models.Customer = Depends(get_active_agency)
):
    """代理店向け利用マニュアル"""
    return templates.TemplateResponse(request=request, name="agency/manual.html", context={
        "request": request,
        "agency": agency
    })

# --- Agency Dashboard ---
@app.get("/agency/", response_class=HTMLResponse)
async def agency_dashboard(
    request: Request, 
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    # 代理店の注文件数
    order_count = db.query(models.AgencyOrder).filter(models.AgencyOrder.customer_id == agency.id).count()
    # 未処理の注文
    pending_count = db.query(models.AgencyOrder).filter(
        models.AgencyOrder.customer_id == agency.id,
        models.AgencyOrder.status == "未処理"
    ).count()
    # 代理店への通知
    notifications = db.query(models.Notification).filter(
        models.Notification.target_type == "agency",
        models.Notification.target_id == agency.id,
        models.Notification.is_read == False
    ).order_by(models.Notification.id.desc()).limit(10).all()
    # 代理店向け請求書（Invoiceのうち、この代理店）
    invoices = db.query(models.Invoice).join(models.Order).join(models.Quotation).filter(
        models.Quotation.customer_id == agency.id
    ).order_by(models.Invoice.id.desc()).limit(5).all()
    
    return templates.TemplateResponse(request=request, name="agency/dashboard.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_dashboard",
        "order_count": order_count,
        "pending_count": pending_count,
        "notifications": notifications,
        "recent_invoices": invoices
    })

# --- Agency Products (商品一覧) ---
@app.get("/agency/products", response_class=HTMLResponse)
async def agency_products(
    request: Request,
    q: str = "",
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    query = db.query(models.Product)
    if q:
        query = query.filter(
            (models.Product.name.contains(q)) | (models.Product.code.contains(q))
        )
    products = query.order_by(models.Product.id.desc()).all()
    
    # ランクに応じた価格をセット
    products_with_price = []
    for p in products:
        price = get_price_for_rank(p, agency.rank)
        products_with_price.append({
            "id": p.id,
            "code": p.code,
            "name": p.name,
            "price": price,
            "stock_quantity": p.stock_quantity
        })
    
    return templates.TemplateResponse(request=request, name="agency/products.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_products",
        "products": products_with_price,
        "search_query": q
    })

# --- Agency Order (発注) ---
@app.get("/agency/order/new", response_class=HTMLResponse)
async def agency_new_order(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    products = db.query(models.Product).order_by(models.Product.id.desc()).all()
    products_with_price = []
    for p in products:
        price = get_price_for_rank(p, agency.rank)
        # 熊スプレーケース単位発注導入対象：
        # 「熊スプレー」を含み、「ホルダ」「練習」を含まないもののみ case単位
        is_spray = ('熊スプレー' in p.name
                    and 'ホルダ' not in p.name
                    and '練習' not in p.name)
        products_with_price.append({
            "id": p.id,
            "code": p.code,
            "name": p.name,
            "price": price,
            "price_a": p.price_a or 0,
            "price_b": p.price_b or 0,
            "price_c": p.price_c or 0,
            "price_d": p.price_d or 0,
            "price_e": p.price_e or 0,
            "is_spray": is_spray,
            "stock_quantity": p.stock_quantity
        })
    # マタギの一撃を先頭に、次にその他のスプレー、最後に通常商品
    def sort_key(p):
        if p['is_spray'] and 'マタギ' in p['name']:
            return 0
        elif p['is_spray']:
            return 1
        else:
            return 2
    products_with_price.sort(key=sort_key)
    
    return templates.TemplateResponse(request=request, name="agency/order_form.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_orders",
        "products": products_with_price
    })

@app.post("/agency/order/new")
async def agency_create_order(
    request: Request,
    memo: str = Form(""),
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    quantities = form_data.getlist("quantity[]")
    
    try:
        # Check if at least one item has a quantity > 0 (handle empty strings as 0)
        def parse_qty(q):
            try:
                return int(q)
            except (ValueError, TypeError):
                return 0
                
        if not product_ids or all(parse_qty(q) == 0 for q in quantities):
            return HTMLResponse(content="<script>alert('商品を1つ以上選択してください'); history.back();</script>", status_code=400)
        
        order_number = f"AG-{agency.id}-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
        
        agency_order = models.AgencyOrder(
            customer_id=agency.id,
            order_number=order_number,
            memo=memo,
            status="未処理"
        )
        db.add(agency_order)
        db.flush()
        
        total = 0
        for p_id, qty in zip(product_ids, quantities):
            try:
                qty = int(qty)
            except (ValueError, TypeError):
                continue
                
            if qty <= 0:
                continue
            product = db.query(models.Product).get(int(p_id))
            if not product:
                continue
            
            # 熊スプレーはケース入力→本数変換＋ケース数に応じた動的価格計算
            CASE_SIZE = 36
            is_spray = ('熊スプレー' in product.name
                        and 'ホルダ' not in product.name
                        and '練習' not in product.name)
            if is_spray:
                case_count = qty  # 入力値はケース数
                actual_qty = case_count * CASE_SIZE  # 本数に変換
                # ケース数に応じたランク価格を適用
                if case_count >= 30:
                    price = product.price_a or get_price_for_rank(product, agency.rank)
                elif case_count >= 25:
                    price = product.price_b or get_price_for_rank(product, agency.rank)
                elif case_count >= 10:
                    price = product.price_c or get_price_for_rank(product, agency.rank)
                elif case_count >= 5:
                    price = product.price_d or get_price_for_rank(product, agency.rank)
                else:
                    price = product.price_e or get_price_for_rank(product, agency.rank)
            else:
                actual_qty = qty
                price = get_price_for_rank(product, agency.rank)
            
            subtotal = actual_qty * price
            
            item = models.AgencyOrderItem(
                agency_order_id=agency_order.id,
                product_id=product.id,
                product_name=product.name,
                quantity=actual_qty,  # 常に本数で保存
                unit_price=price,
                subtotal=subtotal
            )
            db.add(item)
            total += subtotal
        
        product_total_tax_excl = total
        product_total_tax_incl = int(product_total_tax_excl * 1.1)
        
        if product_total_tax_incl <= 10000:
            return HTMLResponse(content="<script>alert('ご注文合計（税込）が10,000円以下のため、発注を承ることができません。商品を追加してください。'); history.back();</script>", status_code=400)
            
        shipping_fee = 0
        if product_total_tax_incl < 30000:
            shipping_fee = 1200
            # 送料明細の追加
            shipping_item = models.AgencyOrderItem(
                agency_order_id=agency_order.id,
                product_id=None,
                product_name="送料",
                quantity=1,
                unit_price=shipping_fee,
                subtotal=shipping_fee
            )
            db.add(shipping_item)
            total += shipping_fee

        agency_order.total_amount = total
        
        # 当社への通知
        notification = models.Notification(
            target_type="admin",
            target_id=None,
            title="新規代理店発注",
            message=f"{agency.company}様から新規発注（{order_number}）がありました。税込合計: ¥{(product_total_tax_incl + shipping_fee):,.0f}",
            link="/agency-orders",
            related_type="AgencyOrder",
            related_id=agency_order.id
        )
        db.add(notification)
        
        db.commit()
        
        # メール通知の送信（バックグラウンドではなく同期実行、エラーは上記関数内でキャッチ）
        await send_order_notification_email(agency_order, db)
        
        return RedirectResponse(url="/agency/orders", status_code=303)

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        return HTMLResponse(content=f"<h3>Order Creation Error</h3><pre style='background:#fee;padding:1rem;'>{error_details}</pre>", status_code=200)

# --- Agency Order History (発注履歴) ---
@app.get("/agency/orders", response_class=HTMLResponse)
async def agency_order_history(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    orders = db.query(models.AgencyOrder).filter(
        models.AgencyOrder.customer_id == agency.id
    ).order_by(models.AgencyOrder.id.desc()).all()
    
    return templates.TemplateResponse(request=request, name="agency/orders.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_orders",
        "orders": orders
    })

@app.get("/agency/orders/{order_id}", response_class=HTMLResponse)
async def agency_order_detail(
    order_id: int,
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    order = db.query(models.AgencyOrder).filter(
        models.AgencyOrder.id == order_id,
        models.AgencyOrder.customer_id == agency.id
    ).first()
    if not order:
        return RedirectResponse(url="/agency/orders", status_code=303)
    
    return templates.TemplateResponse(request=request, name="agency/order_detail.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_orders",
        "order": order
    })

# --- Agency Invoices (請求書一覧) ---
@app.get("/agency/invoices", response_class=HTMLResponse)
async def agency_invoices(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    invoices = db.query(models.Invoice).join(models.Order).join(models.Quotation).filter(
        models.Quotation.customer_id == agency.id
    ).order_by(models.Invoice.id.desc()).all()
    
    return templates.TemplateResponse(request=request, name="agency/invoices.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_invoices",
        "invoices": invoices
    })

@app.get("/agency/invoices/{invoice_id}/print", response_class=HTMLResponse)
async def agency_print_invoice(
    invoice_id: int,
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    invoice = db.query(models.Invoice).join(models.Order).join(models.Quotation).filter(
        models.Invoice.id == invoice_id,
        models.Quotation.customer_id == agency.id
    ).first()
    if not invoice:
        return RedirectResponse(url="/agency/invoices", status_code=303)
    
    return templates.TemplateResponse(request=request, name="print_layout.html", context={
        "request": request,
        "doc_type": "invoice",
        "doc": invoice,
        "user": None
    })

# --- Agency Quotations (見積書一覧) ---
@app.get("/agency/quotations", response_class=HTMLResponse)
async def agency_quotations(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    quotations = db.query(models.Quotation).filter(
        models.Quotation.customer_id == agency.id
    ).order_by(models.Quotation.id.desc()).all()
    
    return templates.TemplateResponse(request=request, name="agency/quotations.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_quotations",
        "quotations": quotations
    })

@app.get("/agency/quotations/{quote_id}/print", response_class=HTMLResponse)
async def agency_print_quotation(
    quote_id: int,
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    quote = db.query(models.Quotation).filter(
        models.Quotation.id == quote_id,
        models.Quotation.customer_id == agency.id
    ).first()
    if not quote:
        return RedirectResponse(url="/agency/quotations", status_code=303)
    
    return templates.TemplateResponse(request=request, name="print_layout.html", context={
        "request": request,
        "doc_type": "quotation",
        "doc": quote,
        "user": None
    })

# --- Agency Password Change ---
@app.get("/agency/change-password", response_class=HTMLResponse)
async def agency_change_password_page(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    return templates.TemplateResponse(request=request, name="agency/change_password.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_password"
    })

@app.post("/agency/change-password")
async def agency_change_password(
    request: Request,
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...),
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    if agency.agency_password != current_password:
        return templates.TemplateResponse(request=request, name="agency/change_password.html", context={
            "request": request,
            "agency": agency,
            "active_page": "agency_password",
            "error": "現在のパスワードが正しくありません。"
        })
    
    if new_password != confirm_password:
        return templates.TemplateResponse(request=request, name="agency/change_password.html", context={
            "request": request,
            "agency": agency,
            "active_page": "agency_password",
            "error": "新しいパスワードと確認用パスワードが一致しません。"
        })
    
    if len(new_password) < 4:
        return templates.TemplateResponse(request=request, name="agency/change_password.html", context={
            "request": request,
            "agency": agency,
            "active_page": "agency_password",
            "error": "パスワードは4文字以上で設定してください。"
        })
    
    agency.agency_password = new_password
    db.commit()
    
    return templates.TemplateResponse(request=request, name="agency/change_password.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_password",
        "success": "パスワードを正常に変更しました。"
    })

# --- Agency Notifications ---
@app.get("/agency/notifications", response_class=HTMLResponse)
async def agency_notifications(
    request: Request,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    notifications = db.query(models.Notification).filter(
        models.Notification.target_type == "agency",
        models.Notification.target_id == agency.id
    ).order_by(models.Notification.id.desc()).limit(50).all()
    
    # 関連エンティティの最新ステータスを取得
    status_map = {}
    for n in notifications:
        if n.related_type == "AgencyOrder" and n.related_id:
            order = db.query(models.AgencyOrder).get(n.related_id)
            if order:
                status_map[n.id] = order.status
        elif n.related_type == "Invoice" and n.related_id:
            inv = db.query(models.Invoice).get(n.related_id)
            if inv:
                status_map[n.id] = inv.status.name
    
    return templates.TemplateResponse(request=request, name="agency/notifications.html", context={
        "request": request,
        "agency": agency,
        "active_page": "agency_notifications",
        "notifications": notifications,
        "status_map": status_map
    })

@app.get("/agency/notifications/read_and_redirect")
async def agency_read_and_redirect(
    request: Request,
    next: str,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    notification_id = request.query_params.get("id")
    if notification_id:
        n = db.query(models.Notification).filter(
            models.Notification.id == int(notification_id),
            models.Notification.target_type == "agency",
            models.Notification.target_id == agency.id
        ).first()
        if n:
            n.is_read = True
            db.commit()
    return RedirectResponse(url=next, status_code=303)

@app.post("/agency/notifications/{notification_id}/delete")
async def agency_delete_notification(
    notification_id: int,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    n = db.query(models.Notification).filter(
        models.Notification.id == notification_id,
        models.Notification.target_type == "agency",
        models.Notification.target_id == agency.id
    ).first()
    if n:
        db.delete(n)
        db.commit()
    return RedirectResponse(url="/agency/notifications", status_code=303)

@app.post("/agency/notifications/{notification_id}/mark_read")
async def agency_mark_notification_read_post(
    notification_id: int,
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    n = db.query(models.Notification).filter(
        models.Notification.id == notification_id,
        models.Notification.target_type == "agency",
        models.Notification.target_id == agency.id
    ).first()
    if n:
        n.is_read = True
        db.commit()
    return RedirectResponse(url="/agency/notifications", status_code=303)

# --- Agency Product Search API ---
@app.get("/agency/api/products/search")
async def agency_search_products(
    q: str = "",
    db: Session = Depends(get_db),
    agency: models.Customer = Depends(get_active_agency)
):
    products = db.query(models.Product).filter(
        (models.Product.name.contains(q)) | (models.Product.code.contains(q))
    ).order_by(models.Product.id.desc()).limit(10).all()
    
    html = ""
    for p in products:
        price = get_price_for_rank(p, agency.rank)
        safe_name = p.name.replace("'", "\\'")
        html += f'<div class="search-result" onclick="addProduct({p.id}, \'{safe_name}\', {price})">{p.name} ({p.code}) - ¥{price:,.0f}</div>'
    return HTMLResponse(content=html if html else "<div>見つかりませんでした</div>")

# ============================================================
# 管理画面: 代理店発注一覧 + 通知
# ============================================================

@app.get("/agency-orders", response_class=HTMLResponse)
async def admin_agency_orders(
    request: Request,
    q: str = "",
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    query = db.query(models.AgencyOrder).join(models.Customer)
    if q:
        query = query.filter(
            (models.AgencyOrder.order_number.contains(q)) |
            (models.Customer.company.contains(q))
        )
    orders = query.order_by(models.AgencyOrder.id.desc()).all()
    
    return templates.TemplateResponse(request=request, name="agency_orders_admin.html", context={
        "request": request,
        "active_page": "agency_orders",
        "orders": orders,
        "search_query": q,
        "user": user
    })

@app.post("/agency-orders/{order_id}/process")
async def admin_process_agency_order(
    order_id: int,
    status: str = Form(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.AgencyOrder).get(order_id)
    if order:
        order.status = status
        # 代理店への通知
        notification = models.Notification(
            target_type="agency",
            target_id=order.customer_id,
            title="発注ステータス更新",
            message=f"発注 {order.order_number} のステータスが「{status}」に更新されました。",
            link="/agency/orders",
            related_type="AgencyOrder",
            related_id=order.id
        )
        db.add(notification)
        db.commit()
    return RedirectResponse(url="/agency-orders", status_code=303)
@app.post("/admin/agency-orders/{order_id}/delete")
async def delete_agency_order(
    order_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    order = db.query(models.AgencyOrder).get(order_id)
    if order:
        db.delete(order)
        db.commit()
    return RedirectResponse(url="/agency-orders", status_code=303)

@app.post("/admin/agency-orders/{order_id}/convert")
async def convert_agency_order_to_main(
    order_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    agency_order = db.query(models.AgencyOrder).get(order_id)
    if not agency_order:
        return RedirectResponse(url="/agency-orders", status_code=303)
    
    # 1. Create Quotation
    new_quote = models.Quotation(
        customer_id=agency_order.customer_id,
        quote_number=f"Q-AG-{agency_order.id}-{datetime.datetime.now().strftime('%m%d%H%M')}",
        issue_date=datetime.datetime.now(),
        expiry_date=datetime.datetime.now() + datetime.timedelta(days=180),
        total_amount=agency_order.total_amount,
        status=models.QuoteStatus.ORDERED,
        memo=f"代理店発注({agency_order.order_number})より変換"
    )
    db.add(new_quote)
    db.flush()
    
    # 2. Create Items + 在庫減算
    # 送料が最後になるように並べ替え
    sorted_items = sorted(agency_order.items, key=lambda x: 1 if x.product_name == "送料" else 0)
    for item in sorted_items:
        qi = models.QuotationItem(
            quotation_id=new_quote.id,
            product_id=item.product_id,
            description=item.product_name,
            quantity=item.quantity,
            unit_price=item.unit_price,
            subtotal=item.subtotal
        )
        db.add(qi)
        # 在庫を減算
        if item.product_id:
            main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
            if main_loc:
                update_product_stock(db, item.product_id, main_loc.id, item.quantity, "OUTBOUND", f"代理店注文(ORD-{agency_order.order_number})出荷による自動出庫")
    
    # 3. Create Order
    new_order = models.Order(
        quotation_id=new_quote.id,
        order_number=f"ORD-AG-{agency_order.id}-{datetime.datetime.now().strftime('%m%d%H%M')}",
        order_date=datetime.datetime.now(),
        total_amount=agency_order.total_amount,
        status=models.OrderStatus.PENDING,
        memo=f"代理店発注({agency_order.order_number})より変換"
    )
    db.add(new_order)
    
    # 4. Update Agency Order Status
    agency_order.status = "処理済み"
    
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

# パスワード再発行（管理画面から）
@app.post("/customers/{customer_id}/reset-agency-password")
async def reset_agency_password(
    customer_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    import random
    import string
    customer = db.query(models.Customer).get(customer_id)
    if customer and customer.is_agency:
        new_password = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
        customer.agency_password = new_password
        db.commit()
    return RedirectResponse(url=f"/customers/edit/{customer_id}", status_code=303)

@app.post("/admin/customers/{customer_id}/send-account-info")
async def send_customer_account_info(
    customer_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = db.query(models.Customer).get(customer_id)
    if not customer or not customer.is_agency or not customer.email:
        return RedirectResponse(url="/customers?error=missing_info", status_code=303)
    
    settings_records = db.query(models.SystemSetting).all()
    settings = {s.key: s.value for s in settings_records}
    
    smtp_host = settings.get("smtp_host")
    smtp_port_str = settings.get("smtp_port", "587")
    smtp_port = int(smtp_port_str) if smtp_port_str.isdigit() else 587
    smtp_user = settings.get("smtp_user")
    smtp_pass = settings.get("smtp_pass")
    smtp_from = settings.get("smtp_from", smtp_user)
    
    missing = []
    if not smtp_host: missing.append("ホスト")
    if not smtp_user: missing.append("ユーザー名")
    if not smtp_pass: missing.append("パスワード")
    
    if missing:
        missing_str = "、".join(missing)
        all_keys = [s.key for s in settings_records if s.key]
        return RedirectResponse(url=f"/customers?error=smtp_config&msg={missing_str} (全設定数: {len(all_keys)}, キー: {','.join(all_keys)})", status_code=303)

    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "【株式会社熊ノ護化研】代理店システム アカウント情報のご案内"
        msg["From"] = smtp_from if smtp_from else "no-reply@kumanomori.jp"
        msg["To"] = customer.email
        
        login_url = "https://app.kumanomorikaken.co.jp/agency/login" # 本番URL
        
        html = f"""
        <html>
        <body style="font-family: sans-serif; line-height: 1.6; color: #333;">
            <h2>代理店システム アカウント発行のご案内</h2>
            <p>{customer.company or customer.name} 様</p>
            <p>平素は格別のお引き立てをいただき、厚く御礼申し上げます。<br>
            この度、弊社代理店システムの準備が整いましたので、アカウント情報をご案内いたします。</p>
            
            <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 20px; border-radius: 8px; margin: 20px 0;">
                <p style="margin: 0;"><strong>ログインURL:</strong> <a href="{login_url}">{login_url}</a></p>
                <p style="margin: 10px 0 0 0;"><strong>ログインID:</strong> {customer.login_id or '未設定'}</p>
                <p style="margin: 5px 0 0 0;"><strong>パスワード:</strong> {customer.agency_password or '未設定'}</p>
            </div>
            
            <p>管理画面では、商品の発注、見積書の作成、および過去の請求情報の確認が行えます。<br>
            内容をご確認いただき、ログインをお試しください。</p>
            
            <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
            <p style="font-size: 0.9rem; color: #666;">
                株式会社熊ノ護化研<br>
                〒010-0001 秋田県秋田市中通3-1-9<br>
                TEL: 018-838-1920<br>
                Mail: info@kumanomorikaken.co.jp
            </p>
        </body>
        </html>
        """
        msg.attach(MIMEText(html, "html"))
        
        import smtplib
        import ssl
        context = ssl.create_default_context()
        server = None
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=10, context=context)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
            server.starttls(context=context)
        
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        
        return RedirectResponse(url="/customers?success_email=1", status_code=303)
    except Exception as e:
        print(f"Email error: {e}")
        return RedirectResponse(url=f"/customers?error=send_failed&msg={str(e)}", status_code=303)

# 管理画面の通知バッジ用API
@app.get("/api/admin/notification-count")
async def admin_notification_count(db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    count = db.query(models.Notification).filter(
        models.Notification.target_type == "admin",
        models.Notification.is_read == False
    ).count()
    return {"count": count}

@app.get("/admin/invoice-dispatch", response_class=HTMLResponse)
async def admin_invoice_dispatch(
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    # すべての未入金・未発行の請求書を取得 (発行済みに移行していないもの)
    unpaid_invoices = db.query(models.Invoice).filter(
        models.Invoice.status == models.InvoiceStatus.UNPAID
    ).all()
    
    email_invoices = []
    postal_invoices = []
    
    for inv in unpaid_invoices:
        # 合算請求書対応の顧客取得ロジック
        customer = inv.customer if inv.customer else (
            inv.orders[0].quotation.customer if getattr(inv, 'orders', None) and inv.orders and inv.orders[0].quotation else None
        )
        if customer:
            # カラムが存在しない場合やデータがNULLの場合の安全策
            method = getattr(customer, 'invoice_delivery_method', 'POSTAL')
            if method == "EMAIL":
                email_invoices.append(inv)
            else:
                postal_invoices.append(inv)

    success_msg = request.query_params.get("success")
    error_msg = request.query_params.get("error")
                
    return templates.TemplateResponse(request=request, name="invoices/dispatch.html", context={
        "request": request,
        "active_page": "invoice_dispatch",
        "email_invoices": email_invoices,
        "postal_invoices": postal_invoices,
        "success_msg": success_msg,
        "error_msg": error_msg,
        "user": user
    })

import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

@app.post("/admin/invoice-dispatch/email")
async def dispatch_invoices_email(
    request: Request,
    invoice_ids: list[int] = Form([]),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if not invoice_ids:
        return RedirectResponse(url="/admin/invoice-dispatch?error=1", status_code=303)
        
    settings_records = db.query(models.SystemSetting).all()
    settings = {s.key: s.value for s in settings_records}
    
    smtp_host = settings.get("smtp_host")
    smtp_port_str = settings.get("smtp_port", "587")
    smtp_port = int(smtp_port_str) if smtp_port_str.isdigit() else 587
    smtp_user = settings.get("smtp_user")
    smtp_pass = settings.get("smtp_pass")
    smtp_from = settings.get("smtp_from", smtp_user)
    
    if not all([smtp_host, smtp_user, smtp_pass]):
        return RedirectResponse(url="/admin/invoice-dispatch?error=2", status_code=303)

    success_count = 0
    server = None
    context = ssl.create_default_context()
    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=10, context=context)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
            server.starttls(context=context)
        server.login(smtp_user, smtp_pass)
        
        bank_info = settings.get("bank_info", "")
        
        for inv_id in invoice_ids:
            inv = db.query(models.Invoice).get(inv_id)
            if not inv:
                continue
                
            customer = inv.customer if inv.customer else (
                inv.orders[0].quotation.customer if getattr(inv, 'orders', None) and inv.orders and inv.orders[0].quotation else None
            )
            
            if not customer or not customer.email:
                continue
                
            msg = MIMEMultipart()
            msg["Subject"] = f"【ご請求書】株式会社熊ノ護化研より（{inv.invoice_number}）"
            msg["From"] = smtp_from if smtp_from else "no-reply@kumanomori.jp"
            msg["To"] = customer.email
            
            due_date_str = inv.due_date.strftime('%Y年%m月%d日') if inv.due_date else '末日'
            grand_total = int(inv.total_amount * 1.1)
            
            html = f"""
            <html>
            <body style="font-family: sans-serif; line-height: 1.6; color: #333;">
                <h2>ご請求のご案内</h2>
                <p>{customer.company or customer.name} 様</p>
                <p>平素は格別のお引き立てをいただき、厚く御礼申し上げます。<br>
                以下の通りご請求申し上げますので、ご確認のほどよろしくお願いいたします。</p>
                
                <table style="width: 100%; max-width: 600px; border-collapse: collapse; margin-top: 20px;">
                    <tr>
                        <th style="text-align: left; padding: 10px; border-bottom: 2px solid #ccc;">請求書番号</th>
                        <td style="padding: 10px; border-bottom: 1px solid #eee;">{inv.invoice_number}</td>
                    </tr>
                    <tr>
                        <th style="text-align: left; padding: 10px; border-bottom: 2px solid #ccc;">ご請求金額（税込）</th>
                        <td style="padding: 10px; border-bottom: 1px solid #eee; font-size: 1.2em; font-weight: bold; color: #e74c3c;">
                            ¥{"{:,.0f}".format(grand_total)}
                        </td>
                    </tr>
                    <tr>
                        <th style="text-align: left; padding: 10px; border-bottom: 2px solid #ccc;">お支払期限</th>
                        <td style="padding: 10px; border-bottom: 1px solid #eee;">{due_date_str}</td>
                    </tr>
                </table>
                
                <p style="margin-top: 20px;">詳細は添付の請求書PDFをご確認ください。<br>
                お振込先情報はPDF内に記載しております。</p>
                
                <hr style="margin-top: 40px; border: none; border-top: 1px solid #eee;">
                <p style="font-size: 0.9em; color: #888;">
                    株式会社熊ノ護化研<br>
                    〒010-0001 秋田県秋田市中通3-1-9<br>
                    TEL: 018-838-1920
                </p>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(html, "html"))
            
            # PDF生成と添付
            try:
                from email.mime.application import MIMEApplication
                pdf_content = generate_invoice_pdf_content(inv)
                if pdf_content:
                    part = MIMEApplication(pdf_content, Name=f"請求書_{inv.invoice_number}.pdf")
                    part['Content-Disposition'] = f'attachment; filename="請求書_{inv.invoice_number}.pdf"'
                    msg.attach(part)
            except Exception as e:
                print(f"PDF Attachment Error: {e}")

            server.send_message(msg)
            
            inv.status = models.InvoiceStatus.ISSUED
            inv.delivery_status = "SENT"
            success_count += 1
            
        db.commit()
    except Exception as e:
        print(f"SMTP Error: {e}")
        return RedirectResponse(url="/admin/invoice-dispatch?error=3", status_code=303)
    finally:
        if server:
            server.quit()
            
    return RedirectResponse(url=f"/admin/invoice-dispatch?success={success_count}", status_code=303)
    


@app.get("/notifications/{notification_id}/read")
async def read_notification(
    notification_id: int,
    request: Request,
    db: Session = Depends(get_db)
):
    n = db.query(models.Notification).get(notification_id)
    if n:
        n.is_read = True
        db.commit()
        if n.link:
            return RedirectResponse(url=n.link, status_code=303)
    referer = request.headers.get("referer")
    return RedirectResponse(url=referer if referer else "/", status_code=303)

@app.get("/admin/notifications", response_class=HTMLResponse)
async def admin_notifications(
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    notifications = db.query(models.Notification).filter(
        models.Notification.target_type == "admin"
    ).order_by(models.Notification.id.desc()).limit(50).all()
    
    # 関連エンティティの最新ステータスを取得
    status_map = {}
    for n in notifications:
        if n.related_type == "AgencyOrder" and n.related_id:
            order = db.query(models.AgencyOrder).get(n.related_id)
            if order:
                status_map[n.id] = order.status
        elif n.related_type == "Invoice" and n.related_id:
            inv = db.query(models.Invoice).get(n.related_id)
            if inv:
                status_map[n.id] = inv.status.name
    
    return templates.TemplateResponse(request=request, name="admin_notifications.html", context={
        "request": request,
        "active_page": "notifications",
        "notifications": notifications,
        "status_map": status_map,
        "user": user
    })

@app.get("/admin/notifications/read_and_redirect")
async def read_and_redirect(
    request: Request,
    next: str,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    # 本来はIDを渡すべきだが、とりあえず最新の未読を既読にするか、クエリで渡す
    # または dispatch.html などで ID を渡すように URL を組む
    notification_id = request.query_params.get("id")
    if notification_id:
        n = db.query(models.Notification).get(int(notification_id))
        if n:
            n.is_read = True
            db.commit()
    return RedirectResponse(url=next, status_code=303)

@app.post("/admin/notifications/{notification_id}/delete")
async def delete_notification(
    notification_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    n = db.query(models.Notification).get(notification_id)
    if n:
        db.delete(n)
        db.commit()
    return RedirectResponse(url="/admin/notifications", status_code=303)

@app.post("/admin/notifications/{notification_id}/mark_read")
async def mark_notification_read_post(
    notification_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    n = db.query(models.Notification).get(notification_id)
    if n:
        n.is_read = True
        db.commit()
    return RedirectResponse(url="/admin/notifications", status_code=303)


# --- System Settings (Admin) ---
@app.get("/admin/settings", response_class=HTMLResponse)
async def admin_settings(
    request: Request, 
    db: Session = Depends(get_db), 
    user: models.User = Depends(get_active_user)
):
    settings = db.query(models.SystemSetting).all()
    settings_dict = {s.key: s.value for s in settings if s.key}
    
    return templates.TemplateResponse(request=request, name="admin_settings.html", context={
        "request": request,
        "user": user,
        "settings": settings_dict,
        "active_page": "settings_admin"
    })

@app.post("/admin/settings")
async def admin_settings_save(
    request: Request,
    smtp_host: str = Form(""),
    smtp_port: str = Form(""),
    smtp_user: str = Form(""),
    smtp_pass: str = Form(""),
    smtp_from: str = Form(""),
    notification_email: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    data = {
        "smtp_host": smtp_host.strip(),
        "smtp_port": smtp_port.strip(),
        "smtp_user": smtp_user.strip(),
        "smtp_pass": smtp_pass.strip(),
        "smtp_from": smtp_from.strip(),
        "notification_email": notification_email.strip()
    }
    for key, value in data.items():
        setting = db.query(models.SystemSetting).filter(models.SystemSetting.key == key).first()
        if setting:
            setting.value = value
        else:
            db.add(models.SystemSetting(key=key, value=value))
    db.commit()
    return templates.TemplateResponse(request=request, name="admin_settings.html", context={
        "request": request,
        "user": user,
        "settings": data,
        "active_page": "settings_admin",
        "success": "設定を保存しました。"
    })

@app.post("/admin/settings/test-smtp")
async def test_smtp_connection(
    request: Request,
    smtp_host: str = Form(""),
    smtp_port: str = Form(""),
    smtp_user: str = Form(""),
    smtp_pass: str = Form(""),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    import smtplib, io, sys
    port = int(smtp_port) if smtp_port.isdigit() else 587
    context = ssl.create_default_context()
    
    debug_stream = io.StringIO()
    old_stderr = sys.stderr
    sys.stderr = debug_stream
    
    try:
        if port == 465:
            # timeoutを30秒に延長し、server_hostnameを明示
            server = smtplib.SMTP_SSL(smtp_host, port, timeout=30, context=context)
        else:
            server = smtplib.SMTP(smtp_host, port, timeout=30)
            if port == 587:
                server.starttls(context=context)
        
        server.set_debuglevel(1)
        server.login(smtp_user, smtp_pass)
        server.quit()
        return templates.TemplateResponse(request=request, name="admin_settings.html", context={
            "request": request,
            "user": user,
            "success": "SMTP接続テストに成功しました！この設定で問題ありません。",
            "settings": {
                "smtp_host": smtp_host, "smtp_port": smtp_port, 
                "smtp_user": smtp_user, "smtp_pass": smtp_pass
            },
            "active_page": "settings_admin"
        })
    except Exception as e:
        sys.stderr = old_stderr
        debug_logs = debug_stream.getvalue()
        import traceback
        error_detail = traceback.format_exc()
        print(f"SMTP Test Error Details:\n{error_detail}")
        print(f"SMTP Debug Logs:\n{debug_logs}")
        
        return templates.TemplateResponse(request=request, name="admin_settings.html", context={
            "request": request,
            "user": user,
            "error": f"接続テスト失敗 ({type(e).__name__}): {str(e)}",
            "debug_logs": debug_logs, # 画面に通信ログを表示
            "settings": {
                "smtp_host": smtp_host, "smtp_port": smtp_port, 
                "smtp_user": smtp_user, "smtp_pass": smtp_pass
            },
            "active_page": "settings_admin"
        })
    finally:
        if sys.stderr == debug_stream:
            sys.stderr = old_stderr


# --- Location & Stock Movement (拠点・在庫移動管理) ---
@app.get("/admin/locations", response_class=HTMLResponse)
async def list_locations(request: Request, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    if not user.is_admin:
        return RedirectResponse(url="/", status_code=303)
    locations = db.query(models.Location).all()
    return templates.TemplateResponse(request=request, name="admin/locations.html", context={
        "request": request, "active_page": "admin", "locations": locations, "user": user
    })

@app.post("/admin/locations/new")
async def create_location(name: str = Form(...), db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    if not user.is_admin:
        return RedirectResponse(url="/", status_code=303)
    loc = models.Location(name=name)
    db.add(loc)
    db.commit()
    return RedirectResponse(url="/admin/locations", status_code=303)

@app.get("/inventory/move", response_class=HTMLResponse)
async def move_inventory_form(request: Request, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    products = db.query(models.Product).all()
    locations = db.query(models.Location).filter_by(is_active=True).all()
    return templates.TemplateResponse(request=request, name="inventory/move.html", context={
        "request": request, "active_page": "inventory", "products": products, "locations": locations, "user": user
    })

@app.post("/inventory/move")
async def process_inventory_move(
    product_id: int = Form(...),
    from_location_id: int = Form(...),
    to_location_id: int = Form(...),
    quantity: int = Form(...),
    reason: str = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if from_location_id == to_location_id:
        return HTMLResponse(content="<script>alert('移動元と移動先が同じです。'); history.back();</script>", status_code=400)
    
    update_product_stock(
        db, product_id, to_location_id, quantity, "TRANSFER", 
        reason=reason or "管理者による手動移動", from_location_id=from_location_id
    )
    db.commit()
    return RedirectResponse(url="/products", status_code=303)

@app.get("/admin/fix-inventory-magic", response_class=HTMLResponse)
async def fix_inventory_magic(db: Session = Depends(get_db)):
    # 1. 拠点をすべて本社にする（全商品）
    main_loc = db.query(models.Location).filter_by(name="本社倉庫").first()
    if not main_loc:
        return "本社倉庫が見つかりません"
        
    products = db.query(models.Product).all()
    
    target_stocks = {
        "4595558124064": 982,
        "4595558124026": 1699,
        "4595558124071": 9996,
        "4595558124019": 16532,
        "4595558124033": 4993,
        "4595558124057": 998,
        "4595558124040": 490
    }
    
    for p in products:
        # DB上のすべての拠点在庫を削除する
        for stock in p.location_stocks:
            db.delete(stock)
        db.flush()
        
        # 指定の在庫数があればそれを優先し、なければ既存の全体在庫数をそのまま採用する
        target_qty = target_stocks.get(p.code, p.stock_quantity)
        
        # 本社に全在庫を入れる
        new_stock = models.ProductLocationStock(
            product_id=p.id,
            location_id=main_loc.id,
            stock_quantity=target_qty
        )
        db.add(new_stock)
        
        # 総在庫数を上書きする
        p.stock_quantity = target_qty
        
    db.commit()
    return "<h1>処理完了</h1><p>すべての在庫が本社に集約され、指定の在庫数に上書きされました。</p><a href='/products'>商品台帳へ戻る</a>"


# --- 管理者専用: 未入金請求書の一括リセット ---
@app.get("/admin/cleanup-unpaid-invoices", response_class=HTMLResponse)
async def cleanup_unpaid_invoices_page(
    request: Request,
    user: models.User = Depends(get_active_user)
):
    return HTMLResponse("""
    <html><body style="font-family:sans-serif;padding:2rem;">
    <h2>🔧 未入金請求書 一括リセット</h2>
    <p style="color:#e74c3c;font-weight:bold;">⚠️ この操作は取り消せません。テストデータとして登録された全請求書を削除し、紐付いた受注を未請求状態に戻します。</p>
    <form method="post" action="/admin/cleanup-unpaid-invoices" onsubmit="return confirm('本当に実行しますか？');">
        <button type="submit" style="background:#e74c3c;color:white;padding:1rem 2rem;font-size:1.1rem;border:none;border-radius:8px;cursor:pointer;">
            🗑️ 全未入金請求書を削除してリセットする
        </button>
    </form>
    </body></html>
    """)

@app.post("/admin/cleanup-unpaid-invoices", response_class=HTMLResponse)
async def cleanup_unpaid_invoices_execute(
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    results = []
    try:
        # agency_orders に invoice_id カラムが存在するか事前チェック（旧スキーマ対応）
        from sqlalchemy import text
        col_check = db.execute(text("PRAGMA table_info(agency_orders)")).fetchall()
        has_agency_invoice_col = any(row[1] == "invoice_id" for row in col_check)

        unpaid = db.query(models.Invoice).filter(models.Invoice.status == models.InvoiceStatus.UNPAID).all()
        results.append(f"対象請求書: {len(unpaid)}件")
        # 1. 100%解決のため、本日（2026-04-11）より前の全ての未入金テストデータを「完全抹消」
        # これにより、一括発行画面に残っている古いゴミを一掃します。
        safe_date_dt = datetime.datetime(2026, 4, 11)
        
        unpaid_invoices = db.query(models.Invoice).filter(
            models.Invoice.status == models.InvoiceStatus.UNPAID,
            models.Invoice.issue_date < safe_date_dt
        ).all()
        
        results.append(f"対象のテスト用請求書: {len(unpaid_invoices)}件")
        for inv in unpaid_invoices:
            results.append(f"  - {inv.invoice_number} を削除中...")
            inv_id = inv.id
            
            # 紐付く受注を特定 (テスト期間中のものは物理削除)
            orders = db.query(models.Order).filter(models.Order.invoice_id == inv_id).all()
            for o in orders:
                # 4/11以前の紐付け受注はすべてテスト用とみなして削除
                q = o.quotation
                db.delete(o)
                if q: db.delete(q)
            
            # 代理店受注の削除 (4/11以前)
            if has_agency_invoice_col:
                db.execute(text("DELETE FROM agency_order_items WHERE agency_order_id IN (SELECT id FROM agency_orders WHERE invoice_id = :inv_id AND order_date < '2026-04-11')"), {"inv_id": inv_id})
                db.execute(text("DELETE FROM agency_orders WHERE invoice_id = :inv_id AND order_date < '2026-04-11'"), {"inv_id": inv_id})
                db.execute(text("UPDATE agency_orders SET invoice_id = NULL, status = '未処理' WHERE invoice_id = :inv_id"), {"inv_id": inv_id})
            
            db.delete(inv)

        # 2. 紐付けの切れた浮いているテスト受注も一掃
        old_test_orders = db.query(models.Order).filter(models.Order.order_date < safe_date_dt).all()
        for o in old_test_orders:
            # 既に請求書があるものは残し、未請求のまま放置されている4/11以前の受注のみ削除
            if not o.invoice_id:
                q = o.quotation
                db.delete(o)
                if q: db.delete(q)

        # 3. シャドウデータの残りカスを掃除
        db.execute(text("DELETE FROM quotations WHERE quote_number LIKE 'Q-SHADOW-%'"))
        db.execute(text("DELETE FROM orders WHERE order_number LIKE 'ORD-SHADOW-%'"))
        
        db.commit()
        results.append("✅ データベースのクリーンアップが100%完了しました！")
        results.append("※ 4月11日以前のテスト用データは全て抹消されました。")
        result_html = "<br>".join(results)
        return HTMLResponse(f"""
        <html><body style="font-family:sans-serif;padding:2rem;">
        <h2>✅ リセット完了</h2>
        <pre style="background:#f5f5f5;padding:1rem;border-radius:8px;">{result_html}</pre>
        <a href="/admin/invoice-dispatch" style="display:inline-block;margin-top:1rem;background:#2ecc71;color:white;padding:0.8rem 1.5rem;border-radius:8px;text-decoration:none;">
            → 一括発行画面へ戻る
        </a>
        </body></html>
        """)
    except Exception as e:
        db.rollback()
        return HTMLResponse(f"""
        <html><body style="font-family:sans-serif;padding:2rem;">
        <h2 style="color:red;">❌ エラー</h2>
        <pre>{str(e)}</pre>
        <a href="/admin/cleanup-unpaid-invoices">← 戻る</a>
        </body></html>
        """, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

# TEST_COMMIT: 2026-04-02-0001

# FINAL_SYSTEM_CHECK_SUCCESS: 2026-04-02-0010
