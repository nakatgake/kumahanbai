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
from email.message import EmailMessage

# Create database tables
models.Base.metadata.create_all(bind=engine)

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
    
    # Multi-Warehouse Migration
    # 1. QuotationItem / AgencyOrderItem location_id
    for item_table in ['quotation_items', 'agency_order_items']:
        cursor.execute(f"PRAGMA table_info({item_table})")
        cols = [row[1] for row in cursor.fetchall()]
        if 'location_id' not in cols:
            print(f"Migrating {item_table}: adding location_id...")
            cursor.execute(f"ALTER TABLE {item_table} ADD COLUMN location_id INTEGER")
            
    conn.commit()
    conn.close()

# Initialize Multi-Warehouse Data
def init_warehouse_data():
    db = SessionLocal()
    try:
        # 1. Ensure at least one location exists
        count = db.query(models.Location).count()
        if count == 0:
            print("No locations found. Initializing '本社'...")
            main_loc = models.Location(name="本社", is_default=True)
            db.add(main_loc)
            db.commit()
            db.refresh(main_loc)
            
            # 2. Migrate existing product stock to this location
            products = db.query(models.Product).all()
            for product in products:
                # Check if stock record exists
                exists = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=main_loc.id).first()
                if not exists:
                    print(f"Migrating stock for {product.name} to {main_loc.name}...")
                    stock = models.ProductStock(
                        product_id=product.id,
                        location_id=main_loc.id,
                        quantity=product.stock_quantity or 0
                    )
                    db.add(stock)
            db.commit()
    except Exception as e:
        print(f"Error initializing warehouse data: {e}")
        db.rollback()
    finally:
        db.close()

migrate_db()
init_warehouse_data()

# --- Security Configuration ---
SECRET_KEY = "kumanogo-secret-key-12345" # 本番環境では環境変数などで管理すべき
pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")
serializer = URLSafeSerializer(SECRET_KEY)

def get_password_hash(password):
    return pwd_context.hash(password)

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

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

class NotAuthenticatedException(Exception):
    pass

app = FastAPI()

@app.exception_handler(NotAuthenticatedException)
async def auth_exception_handler(request: Request, exc: NotAuthenticatedException):
    return RedirectResponse(url="/login")

async def get_active_user(request: Request, db: Session = Depends(get_db)):
    user = await get_current_user(request, db)
    if not user:
        if request.url.path not in ["/login", "/init-admin"] and not request.url.path.startswith("/static"):
            raise NotAuthenticatedException()
    return user

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

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

    stats = {
        "customers": q_customers.count(),
        "products": q_products.count(),
        "quotations": q_quotes.count(),
        "orders": q_orders.count(),
        "invoices": q_invoices.count(),
        "total_sales": q_invoices.filter(models.Invoice.status == models.InvoiceStatus.PAID).with_entities(func.sum(models.Invoice.total_amount)).scalar() or 0,
        "unpaid_total": q_invoices.filter(models.Invoice.status == models.InvoiceStatus.UNPAID).with_entities(func.sum(models.Invoice.total_amount)).scalar() or 0
    }

    return templates.TemplateResponse(request=request, name="dashboard.html", context={
        "request": request,
        "active_page": "dashboard",
        "stats": stats,
        "start_date": start_date,
        "end_date": end_date,
        "user": user
    })

# --- Locations (Warehouses) ---
@app.get("/locations", response_class=HTMLResponse)
async def list_locations(
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="locations/list.html", context={
        "request": request,
        "active_page": "locations",
        "locations": locations,
        "user": user
    })

@app.get("/locations/new", response_class=HTMLResponse)
async def new_location(
    request: Request,
    user: models.User = Depends(get_active_user)
):
    return templates.TemplateResponse(request=request, name="locations/form.html", context={
        "request": request,
        "active_page": "locations",
        "location": None,
        "user": user
    })

@app.post("/locations/new")
async def create_location(
    name: str = Form(...),
    description: str = Form(""),
    is_default: bool = Form(False),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    if is_default:
        # Reset other defaults
        db.query(models.Location).filter(models.Location.is_default == True).update({"is_default": False})

    new_loc = models.Location(
        name=name,
        description=description,
        is_default=is_default
    )
    db.add(new_loc)
    db.commit()
    return RedirectResponse(url="/locations", status_code=303)

@app.get("/locations/edit/{location_id}", response_class=HTMLResponse)
async def edit_location(
    location_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    location = db.query(models.Location).get(location_id)
    if not location:
        return RedirectResponse(url="/locations", status_code=303)
        
    return templates.TemplateResponse(request=request, name="locations/form.html", context={
        "request": request,
        "active_page": "locations",
        "location": location,
        "user": user
    })

@app.post("/locations/edit/{location_id}")
async def update_location(
    location_id: int,
    name: str = Form(...),
    description: str = Form(""),
    is_default: bool = Form(False),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    location = db.query(models.Location).get(location_id)
    if location:
        if is_default and not location.is_default:
            db.query(models.Location).filter(models.Location.is_default == True).update({"is_default": False})
        
        location.name = name
        location.description = description
        location.is_default = is_default
        db.commit()
    return RedirectResponse(url="/locations", status_code=303)

@app.post("/locations/delete/{location_id}")
async def delete_location(
    location_id: int,
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    location = db.query(models.Location).get(location_id)
    if location and not location.is_default:
        # Ensure it has no stocks before deleting. For simplicity with cascades, we just delete or check.
        # It's better to prevent deletion if there is non-zero stock.
        total_stock = sum(s.quantity for s in location.stocks)
        if total_stock == 0:
            db.delete(location)
            db.commit()
    return RedirectResponse(url="/locations", status_code=303)

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
    is_agency: bool = Form(False),
    invoice_delivery_method: str = Form("POSTAL"),
    login_id: Optional[str] = Form(None),
    agency_password: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = models.Customer(
        name=name, company=company, zip_code=zip_code, 
        email=email, phone=phone, address=address, website_url=website_url,
        rank=models.CustomerRank[rank],
        is_agency=is_agency,
        invoice_delivery_method=invoice_delivery_method,
        login_id=login_id if is_agency and login_id else None,
        agency_password=agency_password if is_agency and agency_password else None
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
    is_agency: bool = Form(False),
    invoice_delivery_method: str = Form("POSTAL"),
    login_id: Optional[str] = Form(None),
    agency_password: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = db.query(models.Customer).get(customer_id)
    if customer:
        customer.name = name
        customer.company = company
        customer.zip_code = zip_code
        customer.email = email
        customer.phone = phone
        customer.address = address
        customer.website_url = website_url
        customer.rank = models.CustomerRank[rank]
        customer.is_agency = is_agency
        customer.invoice_delivery_method = invoice_delivery_method
        customer.login_id = login_id if is_agency and login_id else None
        customer.agency_password = agency_password if is_agency and agency_password else None
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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="products/list.html", context={
        "request": request,
        "active_page": "products",
        "products": products,
        "locations": locations,
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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="products/form.html", context={
        "request": request, 
        "active_page": "products",
        "locations": locations,
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
    location_id: int = Form(None),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = models.Product(
        code=code, name=name, 
        unit_price=price_retail, # Keep for compatibility
        price_retail=price_retail,
        price_a=price_a, price_b=price_b, price_c=price_c, price_d=price_d, price_e=price_e,
        stock_quantity=stock_quantity
    )
    db.add(product)
    db.commit()
    db.refresh(product)
    
    # Store initial stock to selected location
    if stock_quantity > 0 and location_id:
        stock = models.ProductStock(product_id=product.id, location_id=location_id, quantity=stock_quantity)
        db.add(stock)
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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="products/form.html", context={
        "request": request, 
        "active_page": "products", 
        "product": product,
        "locations": locations,
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
    location_id: int = Form(None),
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
        
        # 入庫加算: 指定拠点に入庫数を加算する
        if stock_add > 0 and location_id:
            stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=location_id).first()
            if not stock:
                stock = models.ProductStock(product_id=product.id, location_id=location_id, quantity=0)
                db.add(stock)
            stock.quantity += stock_add
            # Update total cache
            product.stock_quantity = sum(s.quantity for s in product.location_stocks) + stock.quantity if not stock.id else sum(s.quantity for s in product.location_stocks)
            product.stock_quantity += stock_add
            
        db.commit()
    return RedirectResponse(url="/products", status_code=303)

@app.post("/products/{product_id}/add_stock")
async def add_stock(
    product_id: int,
    quantity: int = Form(...),
    location_id: int = Form(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    product = db.query(models.Product).get(product_id)
    if product and quantity > 0:
        # Update specific location stock
        stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=location_id).first()
        if not stock:
            stock = models.ProductStock(product_id=product.id, location_id=location_id, quantity=0)
            db.add(stock)
        stock.quantity += quantity
        
        # Update total cache
        product.stock_quantity = sum(s.quantity for s in product.location_stocks) + stock.quantity if not stock.id else sum(s.quantity for s in product.location_stocks)
        # Actually a safer way is just simple addition, but summing ensures consistency.
        # Since the newly added stock object might not be flushed, let's just do:
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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="quotations/form.html", context={
        "request": request,
        "active_page": "quotations",
        "customers": customers,
        "locations": locations,
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
    location_ids = form_data.getlist("location_id[]")
    
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
    for idx, (p_id, p_name, qty, price) in enumerate(zip(product_ids, product_names, quantities, prices)):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        
        # location_id
        loc_id = location_ids[idx] if idx < len(location_ids) and location_ids[idx] else None
        loc_id = int(loc_id) if loc_id else None

        # product_id can be empty for manual entry
        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            location_id=loc_id,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
    
    final_total_tax_excl = int(total * (1 - (discount_rate / 100)))
    if customer_rank != "RETAIL" and int(final_total_tax_excl * 1.1) < 10000:
        return HTMLResponse(content="<script>alert('ご注文税込み金額が1万円以下の為、お受けすることができません。'); history.back();</script>", status_code=400)

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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="quotations/form.html", context={
        "request": request,
        "active_page": "quotations",
        "quotation": quotation,
        "customers": customers,
        "locations": locations,
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
                product = db.query(models.Product).get(old_item.product_id)
                if product:
                    product.stock_quantity += old_item.quantity
                    if old_item.location_id:
                        p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=old_item.location_id).first()
                        if p_stock:
                            p_stock.quantity += old_item.quantity

    # 洗替方式で明細を更新
    db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quote_id).delete()
    
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")
    location_ids = form_data.getlist("location_id[]")

    total = 0
    for idx, (p_id, p_name, qty, price) in enumerate(zip(product_ids, product_names, quantities, prices)):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        
        loc_id = location_ids[idx] if idx < len(location_ids) and location_ids[idx] else None
        loc_id = int(loc_id) if loc_id else None

        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            location_id=loc_id,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        if quotation.status == models.QuoteStatus.ORDERED and pid:
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
                if loc_id:
                    p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=loc_id).first()
                    if not p_stock:
                        p_stock = models.ProductStock(product_id=product.id, location_id=loc_id, quantity=0)
                        db.add(p_stock)
                    p_stock.quantity -= qty
    
    final_total_tax_excl = int(total * (1 - (discount_rate / 100)))
    if customer_rank != "RETAIL" and int(final_total_tax_excl * 1.1) < 10000:
        return HTMLResponse(content="<script>alert('ご注文税込み金額が1万円以下の為、お受けすることができません。'); history.back();</script>", status_code=400)

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
                        product = db.query(models.Product).get(item.product_id)
                        if product:
                            product.stock_quantity += item.quantity
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
                    product = db.query(models.Product).get(item.product_id)
                    if product:
                        product.stock_quantity += item.quantity
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

    # 3. Deduct Stock from Products using specific locations
    for item in quote.items:
        if item.product_id:
            product = db.query(models.Product).get(item.product_id)
            if product:
                product.stock_quantity -= item.quantity
                if item.location_id:
                    p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=item.location_id).first()
                    if not p_stock:
                        p_stock = models.ProductStock(product_id=product.id, location_id=item.location_id, quantity=0)
                        db.add(p_stock)
                    p_stock.quantity -= item.quantity
    
    db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.get("/orders/new", response_class=HTMLResponse)
async def new_order_form(
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customers = db.query(models.Customer).order_by(models.Customer.id.desc()).all()
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    # Generate a default order number
    order_number = f"ORD-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
    return templates.TemplateResponse(request=request, name="orders/form.html", context={
        "request": request,
        "active_page": "orders",
        "customers": customers,
        "locations": locations,
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
    location_ids = form_data.getlist("location_id[]")

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
    for idx, (p_id, p_name, qty, price) in enumerate(zip(product_ids, product_names, quantities, prices)):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        
        loc_id = location_ids[idx] if idx < len(location_ids) and location_ids[idx] else None
        loc_id = int(loc_id) if loc_id else None

        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            location_id=loc_id,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        # Inventory deduction
        if pid:
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
                if loc_id:
                    p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=loc_id).first()
                    if not p_stock:
                        p_stock = models.ProductStock(product_id=product.id, location_id=loc_id, quantity=0)
                        db.add(p_stock)
                    p_stock.quantity -= qty
    
    final_total_tax_excl = int(total * (1 - (discount_rate / 100)))
    if customer_rank != "RETAIL" and int(final_total_tax_excl * 1.1) < 10000:
        return HTMLResponse(content="<script>alert('ご注文税込み金額が1万円以下の為、お受けすることができません。'); history.back();</script>", status_code=400)

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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    return templates.TemplateResponse(request=request, name="orders/form.html", context={
        "request": request,
        "active_page": "orders",
        "order": order,
        "customers": customers,
        "locations": locations,
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
            product = db.query(models.Product).get(old_item.product_id)
            if product:
                product.stock_quantity += old_item.quantity
                if old_item.location_id:
                    p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=old_item.location_id).first()
                    if p_stock:
                        p_stock.quantity += old_item.quantity

    # 洗替方式で明細を更新
    db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quotation.id).delete()
    
    form_data = await request.form()
    product_ids = form_data.getlist("product_id[]")
    product_names = form_data.getlist("product_name[]")
    quantities = form_data.getlist("quantity[]")
    prices = form_data.getlist("price[]")
    location_ids = form_data.getlist("location_id[]")

    total = 0
    for idx, (p_id, p_name, qty, price) in enumerate(zip(product_ids, product_names, quantities, prices)):
        qty = int(qty)
        price = float(price)
        subtotal = qty * price
        
        loc_id = location_ids[idx] if idx < len(location_ids) and location_ids[idx] else None
        loc_id = int(loc_id) if loc_id else None

        pid = int(p_id) if p_id and p_id != "" else None
        
        item = models.QuotationItem(
            quotation_id=quotation.id,
            product_id=pid,
            location_id=loc_id,
            description=p_name,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )
        db.add(item)
        total += subtotal
        
        if pid:
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
                if loc_id:
                    p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=loc_id).first()
                    if not p_stock:
                        p_stock = models.ProductStock(product_id=product.id, location_id=loc_id, quantity=0)
                        db.add(p_stock)
                    p_stock.quantity -= qty
    
    final_total_tax_excl = int(total * (1 - (discount_rate / 100)))
    if customer_rank != "RETAIL" and int(final_total_tax_excl * 1.1) < 10000:
        return HTMLResponse(content="<script>alert('ご注文税込み金額が1万円以下の為、お受けすることができません。'); history.back();</script>", status_code=400)

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
        order.invoice.due_date = order.order_date + datetime.timedelta(days=30)
        
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
                        product = db.query(models.Product).get(item.product_id)
                        if product:
                            product.stock_quantity += item.quantity
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
    query = db.query(models.Invoice).join(models.Order).join(models.Quotation).join(models.Customer)
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
    query = db.query(models.Invoice).join(models.Order).join(models.Quotation).join(models.Customer)
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

    headers = ["請求番号", "顧客名", "発行日", "支払期限", "請求金額(税抜)"]
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)
        worksheet.set_column(i, i, 20)

    for row, inv in enumerate(invoices, 1):
        worksheet.write(row, 0, inv.invoice_number, border_fmt)
        worksheet.write(row, 1, inv.order.quotation.customer.company, border_fmt)
        worksheet.write(row, 2, inv.issue_date, date_fmt)
        worksheet.write(row, 3, inv.due_date, date_fmt)
        worksheet.write(row, 4, inv.total_amount, num_fmt)

    workbook.close()
    output.seek(0)
    
    filename = f"invoices_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/orders/{order_id}/invoice")
async def create_invoice(order_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    order = db.query(models.Order).get(order_id)
    if not order or order.invoice:
        return RedirectResponse(url="/orders", status_code=303)

    # Create Invoice
    invoice = models.Invoice(
        order_id=order.id,
        invoice_number=order.order_number.replace("ORD-", "INV-"),
        issue_date=order.order_date, # Default to order date
        due_date=order.order_date + datetime.timedelta(days=30),
        total_amount=order.total_amount,
        discount_rate=order.discount_rate,
        is_bulk_discount=order.is_bulk_discount,
        status=models.InvoiceStatus.UNPAID,
        memo=order.memo
    )
    db.add(invoice)
    
    # Update order status to SHIPPED (出荷済み)
    order.status = models.OrderStatus.SHIPPED
    db.commit()
    return RedirectResponse(url="/orders", status_code=303)

@app.post("/orders/{order_id}/cancel_shipping")
async def cancel_shipping(order_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    order = db.query(models.Order).get(order_id)
    if not order or order.status != models.OrderStatus.SHIPPED:
        return RedirectResponse(url="/orders", status_code=303)
    
    # Delete associated invoice
    if order.invoice:
        db.delete(order.invoice)
    
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

@app.post("/invoices/bulk_print")
async def bulk_print_invoices(request: Request, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    form_data = await request.form()
    invoice_ids = form_data.getlist("invoice_ids[]")
    if not invoice_ids:
        return RedirectResponse(url="/invoices", status_code=303)
    
    # 印刷対象は「未発行」のみ（請求書発行済・入金済はスキップ）
    invoices = db.query(models.Invoice).filter(
        models.Invoice.id.in_(invoice_ids),
        models.Invoice.status == models.InvoiceStatus.UNPAID
    ).all()
    
    if not invoices:
        return RedirectResponse(url="/invoices", status_code=303)
    
    # 印刷した分を「請求書発行済」に自動変更
    for inv in invoices:
        inv.status = models.InvoiceStatus.ISSUED
    db.commit()
    
    return templates.TemplateResponse(request=request, name="invoices/bulk_print.html", context={
        "request": request,
        "invoices": invoices,
        "user": user
    })

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
    invoice.discount_rate = discount_rate
    invoice.is_bulk_discount = is_bulk_discount
    invoice.memo = memo
    
    db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.post("/invoices/delete/{invoice_id}")
async def delete_invoice(invoice_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    if invoice:
        # 請求書に紐づく受注の明細から在庫を戻す
        order = invoice.order
        if order and order.quotation:
            for item in order.quotation.items:
                if item.product_id:
                    product = db.query(models.Product).get(item.product_id)
                    if product:
                        product.stock_quantity += item.quantity
        # 紐づく受注も削除（受注がなくなれば請求も不要）
        if order:
            db.delete(order)
        db.delete(invoice)
        db.commit()
    response = RedirectResponse(url="/invoices", status_code=303)
    response.headers["HX-Refresh"] = "true"
    return response

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
    try:
        settings = db.query(models.SystemSetting).all()
        s = {s.key: s.value for s in settings}
        
        # 宛先: 設定がなければ info@kumanomorikaken.co.jp に送る
        target = s.get("notification_email") or "info@kumanomorikaken.co.jp"
        
        if not s.get("smtp_host"):
            print("Email notification skipped: no SMTP host configured")
            return

        msg = EmailMessage()
        company_name = order.customer.company if order.customer else "不明な代理店"
        content = f"""代理店：{company_name} 様より新規発注がありました。

【受注番号】: {order.order_number}
【発注日時】: {order.order_date.strftime('%Y/%m/%d %H:%M') if order.order_date else '-'}
【合計金額】: ¥{'{:,.0f}'.format(order.total_amount)} (税抜)

詳細は管理画面の「代理店発注」ページ、または以下の通知一覧よりご確認ください。
https://app.kumanomorikaken.co.jp/admin/notifications
"""
        msg.set_content(content)
        msg['Subject'] = f"【代理店サイト】新規発注のお知らせ ({company_name}様)"
        msg['From'] = s.get("smtp_from")
        msg['To'] = target

        # SMTP configuration
        smtp_host = s.get("smtp_host")
        smtp_port_val = s.get("smtp_port") or "587"
        try:
            smtp_port = int(smtp_port_val)
        except ValueError:
            smtp_port = 587
            
        smtp_user = s.get("smtp_user")
        smtp_pass = s.get("smtp_pass")

        # SMTP Connection
        context = ssl.create_default_context()
        if smtp_port == 465:
            # timeoutを30秒に延長し、server_hostnameを明示
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
        print(f"Email sent to {target}")
    except Exception as e:
        print(f"Failed to send email: {e}")
        # Note: We don't raise the error here to avoid breaking the order creation flow

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
        products_with_price.append({
            "id": p.id,
            "code": p.code,
            "name": p.name,
            "price": price,
            "stock_quantity": p.stock_quantity
        })
    
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
            price = get_price_for_rank(product, agency.rank)
            subtotal = qty * price
            
            item = models.AgencyOrderItem(
                agency_order_id=agency_order.id,
                product_id=product.id,
                product_name=product.name,
                quantity=qty,
                unit_price=price,
                subtotal=subtotal
            )
            db.add(item)
            total += subtotal
        
        agency_order.total_amount = total
        
        # 当社への通知
        notification = models.Notification(
            target_type="admin",
            target_id=None,
            title="新規代理店発注",
            message=f"{agency.company}様から新規発注（{order_number}）がありました。合計: ¥{total:,.0f}",
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
    locations = db.query(models.Location).order_by(models.Location.is_default.desc(), models.Location.name).all()
    
    return templates.TemplateResponse(request=request, name="agency_orders_admin.html", context={
        "request": request,
        "active_page": "agency_orders",
        "orders": orders,
        "locations": locations,
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
    location_id: int = Form(...),
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
    for item in agency_order.items:
        qi = models.QuotationItem(
            quotation_id=new_quote.id,
            product_id=item.product_id,
            location_id=location_id,
            description=item.product_name,
            quantity=item.quantity,
            unit_price=item.unit_price,
            subtotal=item.subtotal
        )
        db.add(qi)
        # 在庫を減算
        if item.product_id:
            product = db.query(models.Product).get(item.product_id)
            if product:
                product.stock_quantity -= item.quantity
                
                p_stock = db.query(models.ProductStock).filter_by(product_id=product.id, location_id=location_id).first()
                if not p_stock:
                    p_stock = models.ProductStock(product_id=product.id, location_id=location_id, quantity=0)
                    db.add(p_stock)
                p_stock.quantity -= item.quantity
    
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
        # 通常の注文(Order)にはQuotationを介してCustomerが紐付いている
        if inv.order and inv.order.quotation and inv.order.quotation.customer:
            customer = inv.order.quotation.customer
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
            if not inv or not inv.order or not inv.order.quotation or not inv.order.quotation.customer:
                continue
            
            customer = inv.order.quotation.customer
            if not customer.email:
                continue
                
            msg = MIMEMultipart("alternative")
            msg["Subject"] = f"【株式会社熊ノ護化研】ご請求書（{inv.invoice_number}）のご案内"
            msg["From"] = smtp_from if smtp_from else "no-reply@kumanomori.jp"
            msg["To"] = customer.email
            
            due_date_str = inv.due_date.strftime('%Y年%m月%d日') if inv.due_date else '末日'
            bank_html = bank_info.replace(chr(10), '<br>') if bank_info else '※銀行振込先は別途ご案内いたします。'
            
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
                        <th style="text-align: left; padding: 10px; border-bottom: 2px solid #ccc;">ご請求金額</th>
                        <td style="padding: 10px; border-bottom: 1px solid #eee; font-size: 1.2em; font-weight: bold; color: #e74c3c;">
                            ¥{"{:,.0f}".format(inv.total_amount)}
                        </td>
                    </tr>
                    <tr>
                        <th style="text-align: left; padding: 10px; border-bottom: 2px solid #ccc;">お支払期限</th>
                        <td style="padding: 10px; border-bottom: 1px solid #eee;">{due_date_str}</td>
                    </tr>
                </table>
                
                <h3 style="margin-top: 30px;">■ お振込先</h3>
                <div style="background: #f8f9fa; padding: 15px; border-radius: 5px;">
                    {bank_html}
                </div>
                
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
            server.send_message(msg)
            
            inv.status = models.InvoiceStatus.ISSUED
            success_count += 1
            
        db.commit()
    except Exception as e:
        print(f"SMTP Error: {e}")
        return RedirectResponse(url="/admin/invoice-dispatch?error=3", status_code=303)
    finally:
        if server:
            server.quit()
            
    return RedirectResponse(url=f"/admin/invoice-dispatch?success={success_count}", status_code=303)
    
@app.get("/invoices/bulk-print", response_class=HTMLResponse)
async def admin_bulk_print_invoices(
    request: Request,
    invoice_ids: list[int] = Query(...),
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    # 印刷対象の請求書を取得
    invoices = db.query(models.Invoice).filter(
        models.Invoice.id.in_(invoice_ids)
    ).all()
    
    if not invoices:
        return RedirectResponse(url="/admin/invoice-dispatch", status_code=303)
    
    # 印刷した分を「請求書発行済」に自動変更
    for inv in invoices:
        if inv.status == models.InvoiceStatus.UNPAID:
            inv.status = models.InvoiceStatus.ISSUED
    db.commit()
    
    return templates.TemplateResponse(request=request, name="invoices/bulk_print.html", context={
        "request": request,
        "invoices": invoices,
        "user": user
    })

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

# ============================================================
# 月次請求書自動発行スクリプト（手動実行とスケジュール用エンドポイント）
# ============================================================
@app.post("/admin/generate-monthly-invoices")
async def generate_monthly_invoices(
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    """前月の代理店発注をまとめて請求書を生成する"""
    now = datetime.datetime.now()
    # 前月の期間
    first_of_current = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_month_end = first_of_current - datetime.timedelta(seconds=1)
    last_month_start = last_month_end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
    # 代理店一覧
    agencies = db.query(models.Customer).filter(models.Customer.is_agency == True).all()
    generated_count = 0
    
    for agency in agencies:
        # 前月の処理済み発注を集計
        orders = db.query(models.AgencyOrder).filter(
            models.AgencyOrder.customer_id == agency.id,
            models.AgencyOrder.status == "処理済み",
            models.AgencyOrder.order_date >= last_month_start,
            models.AgencyOrder.order_date <= last_month_end
        ).all()
        
        if not orders:
            continue
        
        total = sum(o.total_amount for o in orders)
        month_str = last_month_start.strftime('%Y%m')
        inv_number = f"AGINV-{agency.id}-{month_str}"
        
        # 既に同月の請求書があったらスキップ
        existing = db.query(models.Invoice).filter(models.Invoice.invoice_number == inv_number).first()
        if existing:
            continue
        
        # まず代理店発注用の見積＋受注をシャドウ作成
        shadow_quote = models.Quotation(
            customer_id=agency.id,
            quote_number=f"Q-{inv_number}",
            issue_date=first_of_current,
            expiry_date=first_of_current + datetime.timedelta(days=30),
            total_amount=total,
            status=models.QuoteStatus.ORDERED,
            memo=f"{last_month_start.strftime('%Y年%m月')}分 代理店月次請求"
        )
        db.add(shadow_quote)
        db.flush()
        
        # 明細をまとめる
        for order in orders:
            for item in order.items:
                qi = models.QuotationItem(
                    quotation_id=shadow_quote.id,
                    product_id=item.product_id,
                    description=f"[{order.order_number}] {item.product_name}",
                    quantity=item.quantity,
                    unit_price=item.unit_price,
                    subtotal=item.subtotal
                )
                db.add(qi)
        
        shadow_order = models.Order(
            quotation_id=shadow_quote.id,
            order_number=f"ORD-{inv_number}",
            order_date=first_of_current,
            total_amount=total,
            status=models.OrderStatus.COMPLETED,
            memo=f"{last_month_start.strftime('%Y年%m月')}分 代理店月次請求"
        )
        db.add(shadow_order)
        db.flush()
        
        invoice = models.Invoice(
            order_id=shadow_order.id,
            invoice_number=inv_number,
            issue_date=first_of_current + datetime.timedelta(days=1),  # 翌月2日
            due_date=first_of_current + datetime.timedelta(days=30),
            total_amount=total,
            status=models.InvoiceStatus.UNPAID,
            memo=f"{last_month_start.strftime('%Y年%m月')}分 月次集計請求書"
        )
        db.add(invoice)
        
        # 代理店への通知
        notification = models.Notification(
            target_type="agency",
            target_id=agency.id,
            title="月次請求書発行",
            message=f"{last_month_start.strftime('%Y年%m月')}分の請求書（{inv_number}）が発行されました。金額: ¥{total:,.0f}",
            link="/agency/invoices",
            related_type="Invoice",
            related_id=invoice.id
        )
        db.add(notification)
        generated_count += 1
    
    db.commit()
    return RedirectResponse(url=f"/invoices?success_gen={generated_count}", status_code=303)

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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
