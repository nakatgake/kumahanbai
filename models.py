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
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    quotations = relationship("Quotation", back_populates="customer", cascade="all, delete-orphan")

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
    stock_quantity = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

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

    customer = relationship("Customer", back_populates="quotations")
    items = relationship("QuotationItem", back_populates="quotation", cascade="all, delete-orphan")
    order = relationship("Order", back_populates="quotation", uselist=False, cascade="all, delete-orphan")

class QuotationItem(Base):
    __tablename__ = "quotation_items"
    id = Column(Integer, primary_key=True, index=True)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="CASCADE"))
    product_id = Column(Integer, ForeignKey("products.id", ondelete="SET NULL"))
    description = Column(String)
    quantity = Column(Integer)
    unit_price = Column(Float)
    subtotal = Column(Float)

    quotation = relationship("Quotation", back_populates="items")
    product = relationship("Product")

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
    status = Column(Enum(InvoiceStatus), default=InvoiceStatus.UNPAID)

    order = relationship("Order", back_populates="invoice")

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    full_name = Column(String)
    is_active = Column(Boolean, default=True)
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)
