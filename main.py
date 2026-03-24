from fastapi import FastAPI, Depends, Request, Form, File, UploadFile
from typing import Optional
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
import shutil
import os
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
# --- Customers ---
@app.get("/customers", response_class=HTMLResponse)
async def list_customers(
    request: Request, 
    q: str = "", 
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
    customers = query.all()
    return templates.TemplateResponse(request=request, name="customers/list.html", context={
        "request": request,
        "active_page": "customers",
        "customers": customers,
        "search_query": q,
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
    customers = query.all()
    
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
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customer = models.Customer(
        name=name, company=company, zip_code=zip_code, 
        email=email, phone=phone, address=address, website_url=website_url,
        rank=models.CustomerRank[rank]
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
    products = query.all()
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
    products = query.all()
    
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
        unit_price=price_retail, # Keep for compatibility
        price_retail=price_retail,
        price_a=price_a, price_b=price_b, price_c=price_c, price_d=price_d, price_e=price_e,
        stock_quantity=stock_quantity
    )
    db.add(product)
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
        product.stock_quantity = stock_quantity
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
        
    quotations = query.all()
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
    
    quotations = query.all()
    
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
    customers = db.query(models.Customer).all()
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
        status=models.QuoteStatus.DRAFT
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
    
    quotation.total_amount = total * (1 - (discount_rate / 100))
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
    customers = db.query(models.Customer).all()
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

    if quotation.status == models.QuoteStatus.ORDERED:
        for old_item in db.query(models.QuotationItem).filter(models.QuotationItem.quotation_id == quotation.id).all():
            if old_item.product_id:
                product = db.query(models.Product).get(old_item.product_id)
                if product:
                    product.stock_quantity += old_item.quantity

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
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
    
    quotation.total_amount = total * (1 - (discount_rate / 100))
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
        status=models.QuoteStatus.DRAFT
    )
    db.add(new_quote)
    db.flush()
    
    # Copy items
    for item in original.items:
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
        
    orders = query.all()
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
        
    orders = query.all()
    
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
        status=models.OrderStatus.PENDING
    )
    db.add(order)

    # 2. Update Quotation Status
    quote.status = models.QuoteStatus.ORDERED

    # 3. Deduct Stock from Products
    for item in quote.items:
        product = db.query(models.Product).get(item.product_id)
        if product:
            product.stock_quantity -= item.quantity
    
    db.commit()
    return RedirectResponse(url="/quotations", status_code=303)

@app.get("/orders/new", response_class=HTMLResponse)
async def new_order_form(
    request: Request, 
    db: Session = Depends(get_db),
    user: models.User = Depends(get_active_user)
):
    customers = db.query(models.Customer).all()
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
        total_amount=0
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
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
    
    quotation.total_amount = total * (1 - (discount_rate / 100))
    
    order = models.Order(
        quotation_id=quotation.id,
        order_number=order_number,
        order_date=datetime.datetime.strptime(order_date, '%Y-%m-%d'),
        total_amount=total * (1 - (discount_rate / 100)),
        discount_rate=discount_rate,
        is_bulk_discount=is_bulk_discount,
        status=models.OrderStatus.PENDING
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
    customers = db.query(models.Customer).all()
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
            product = db.query(models.Product).get(pid)
            if product:
                product.stock_quantity -= qty
    
    quotation.total_amount = total * (1 - (discount_rate / 100))
    order.total_amount = quotation.total_amount
    order.discount_rate = discount_rate
    order.is_bulk_discount = is_bulk_discount
    quotation.discount_rate = discount_rate
    quotation.is_bulk_discount = is_bulk_discount
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
        status=models.QuoteStatus.ORDERED
    )
    db.add(new_quote)
    db.flush()
    
    for item in original_quote.items:
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
        status=models.OrderStatus.PENDING
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
        
    invoices = query.all()
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
        
    invoices = query.all()
    
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
        status=models.InvoiceStatus.UNPAID
    )
    db.add(invoice)
    
    # Update order status to SHIPPED (出荷済み)
    order.status = models.OrderStatus.SHIPPED
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
    
    db.commit()
    return RedirectResponse(url="/invoices", status_code=303)

@app.post("/invoices/delete/{invoice_id}")
async def delete_invoice(invoice_id: int, db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    invoice = db.query(models.Invoice).get(invoice_id)
    if invoice:
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
    ).limit(10).all()
    # Simple HTML response for HTMX
    html = ""
    for p in products:
        # p.name をエスケープ (JS 用)
        safe_name = p.name.replace("'", "\\'")
        # prices も安全に文字列化
        prices_js = f"{{retail: {p.price_retail}, a: {p.price_a}, b: {p.price_b}, c: {p.price_c}, d: {p.price_d}, e: {p.price_e}}}"
        html += f'<div class="search-result" onclick="selectProduct({p.id}, \'{safe_name}\', {prices_js})">{p.name} ({p.code}) - ¥{p.price_retail:,.0f}</div>'
    return HTMLResponse(content=html if html else "<div>見つかりませんでした</div>")

# API for HTMX Customer Search
@app.get("/api/customers/search")
async def search_customers(q: str = "", db: Session = Depends(get_db), user: models.User = Depends(get_active_user)):
    customers = db.query(models.Customer).filter(
        (models.Customer.name.contains(q)) | (models.Customer.company.contains(q))
    ).limit(10).all()
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
    users = db.query(models.User).all()
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
        users = db.query(models.User).all()
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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
