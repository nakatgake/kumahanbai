from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey, Enum, Boolean
from sqlalchemy.orm import relationship
from database import Base
import datetime
import enum

class QuoteStatus(enum.Enum):
    DRAFT = "下書き"
    SENT = "送付済み"
    ORDERED = "受注済み"
    CANCELLED = "失注"

class OrderStatus(enum.Enum):
    PENDING = "未出荷"
    SHIPPED = "出荷済み"
    COMPLETED = "完了"

class InvoiceStatus(enum.Enum):
    UNPAID = "未入金"
    PAID = "入金済み"
    ISSUED = "発行済み"

class CustomerRank(enum.Enum):
    RETAIL = "小売り"
    RANK_A = "Aランク"
    RANK_B = "Bランク"
    RANK_C = "Cランク"
    RANK_D = "Dランク"
    RANK_E = "Eランク"

class Customer(Base):
    __tablename__ = "customers"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True, nullable=True)
    company = Column(String)
    zip_code = Column(String)
    email = Column(String)
    phone = Column(String)
    address = Column(String)
    website_url = Column(String)
    rank = Column(Enum(CustomerRank), default=CustomerRank.RETAIL)
    # Agency fields
    is_agency = Column(Boolean, default=False)
    login_id = Column(String, unique=True, nullable=True)
    agency_password = Column(String, nullable=True)  # 平文で保存（当社が確認可能）
    invoice_delivery_method = Column(String, default="POSTAL") # "POSTAL" or "EMAIL"
    closing_day = Column(Integer, nullable=True) # 1-31 (31 is end of month)
    payment_term_months = Column(Integer, default=1) # 0:Same month, 1:Next month, 2:Month after next
    payment_day = Column(Integer, nullable=True) # 1-31 (31 is end of month)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    quotations = relationship("Quotation", back_populates="customer", cascade="all, delete-orphan")
    agency_orders = relationship("AgencyOrder", back_populates="customer", cascade="all, delete-orphan")

class Product(Base):
    __tablename__ = "products"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    name = Column(String, index=True)
    unit_price = Column(Float) # Legacy field, will map to price_retail
    price_retail = Column(Float, default=0.0)
    price_a = Column(Float, default=0.0)
    price_b = Column(Float, default=0.0)
    price_c = Column(Float, default=0.0)
    price_d = Column(Float, default=0.0)
    price_e = Column(Float, default=0.0)
    stock_quantity = Column(Integer, default=0) # 全拠点合計数 (キャッシュ)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    location_stocks = relationship("ProductStock", back_populates="product", cascade="all, delete-orphan")

class Location(Base):
    __tablename__ = "locations"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True)
    description = Column(String, nullable=True)
    is_default = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    stocks = relationship("ProductStock", back_populates="location", cascade="all, delete-orphan")

class ProductStock(Base):
    __tablename__ = "product_stocks"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id", ondelete="CASCADE"))
    location_id = Column(Integer, ForeignKey("locations.id", ondelete="CASCADE"))
    quantity = Column(Integer, default=0)

    product = relationship("Product", back_populates="location_stocks")
    location = relationship("Location", back_populates="stocks")

class Quotation(Base):
    __tablename__ = "quotations"
    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(Integer, ForeignKey("customers.id"))
    quote_number = Column(String, unique=True)
    issue_date = Column(DateTime, default=datetime.datetime.utcnow)
    expiry_date = Column(DateTime)
    payment_due_date = Column(DateTime)
    payment_method = Column(String, default="銀行振り込み")
    total_amount = Column(Float, default=0.0)
    discount_rate = Column(Float, default=0.0)
    is_bulk_discount = Column(Boolean, default=False)
    status = Column(Enum(QuoteStatus), default=QuoteStatus.DRAFT)
    memo = Column(String)

    customer = relationship("Customer", back_populates="quotations")
    items = relationship("QuotationItem", back_populates="quotation", cascade="all, delete-orphan")
    order = relationship("Order", back_populates="quotation", uselist=False, cascade="all, delete-orphan")

class QuotationItem(Base):
    __tablename__ = "quotation_items"
    id = Column(Integer, primary_key=True, index=True)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="CASCADE"))
    product_id = Column(Integer, ForeignKey("products.id", ondelete="SET NULL"))
    location_id = Column(Integer, ForeignKey("locations.id", ondelete="SET NULL"), nullable=True)
    description = Column(String)
    quantity = Column(Integer)
    unit_price = Column(Float)
    subtotal = Column(Float)

    quotation = relationship("Quotation", back_populates="items")
    product = relationship("Product")
    location = relationship("Location")

class Order(Base):
    __tablename__ = "orders"
    id = Column(Integer, primary_key=True, index=True)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="CASCADE"))
    order_number = Column(String, unique=True)
    order_date = Column(DateTime, default=datetime.datetime.utcnow)
    total_amount = Column(Float)
    discount_rate = Column(Float, default=0.0)
    is_bulk_discount = Column(Boolean, default=False)
    status = Column(Enum(OrderStatus), default=OrderStatus.PENDING)
    memo = Column(String)

    quotation = relationship("Quotation", back_populates="order")
    invoice = relationship("Invoice", back_populates="order", uselist=False, cascade="all, delete-orphan")

class Invoice(Base):
    __tablename__ = "invoices"
    id = Column(Integer, primary_key=True, index=True)
    order_id = Column(Integer, ForeignKey("orders.id", ondelete="CASCADE"))
    invoice_number = Column(String, unique=True)
    issue_date = Column(DateTime, default=datetime.datetime.utcnow)
    due_date = Column(DateTime)
    total_amount = Column(Float)
    discount_rate = Column(Float, default=0.0)
    is_bulk_discount = Column(Boolean, default=False)
    status = Column(Enum(InvoiceStatus), default=InvoiceStatus.UNPAID)
    memo = Column(String)

    order = relationship("Order", back_populates="invoice")

class AgencyOrder(Base):
    """代理店からの発注"""
    __tablename__ = "agency_orders"
    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(Integer, ForeignKey("customers.id"))
    order_number = Column(String, unique=True)
    order_date = Column(DateTime, default=datetime.datetime.utcnow)
    total_amount = Column(Float, default=0.0)
    status = Column(String, default="未処理")  # 未処理 / 処理済み / キャンセル
    memo = Column(String, nullable=True)

    customer = relationship("Customer", back_populates="agency_orders")
    items = relationship("AgencyOrderItem", back_populates="agency_order", cascade="all, delete-orphan")

class AgencyOrderItem(Base):
    """代理店発注の明細"""
    __tablename__ = "agency_order_items"
    id = Column(Integer, primary_key=True, index=True)
    agency_order_id = Column(Integer, ForeignKey("agency_orders.id", ondelete="CASCADE"))
    product_id = Column(Integer, ForeignKey("products.id", ondelete="SET NULL"), nullable=True)
    location_id = Column(Integer, ForeignKey("locations.id", ondelete="SET NULL"), nullable=True)
    product_name = Column(String)
    quantity = Column(Integer)
    unit_price = Column(Float)
    subtotal = Column(Float)

    agency_order = relationship("AgencyOrder", back_populates="items")
    product = relationship("Product")
    location = relationship("Location")

class Notification(Base):
    """通知"""
    __tablename__ = "notifications"
    id = Column(Integer, primary_key=True, index=True)
    target_type = Column(String)  # "admin" or "agency"
    target_id = Column(Integer, nullable=True)  # customer_id for agency, null for admin
    title = Column(String)
    message = Column(String)
    is_read = Column(Boolean, default=False)
    link = Column(String, nullable=True)
    related_type = Column(String, nullable=True)  # "AgencyOrder", "Invoice", etc.
    related_id = Column(Integer, nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    full_name = Column(String)
    is_active = Column(Boolean, default=True)
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

class SystemSetting(Base):
    __tablename__ = "system_settings"
    id = Column(Integer, primary_key=True, index=True)
    key = Column(String, unique=True, index=True)
    value = Column(String)
