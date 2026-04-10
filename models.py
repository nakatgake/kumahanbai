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
    honorific = Column(String, default="御中")
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
    stock_quantity = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    location_stocks = relationship("ProductLocationStock", back_populates="product", cascade="all, delete-orphan")

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
    invoice_id = Column(Integer, ForeignKey("invoices.id", ondelete="SET NULL"), nullable=True)
    memo = Column(String)

    quotation = relationship("Quotation", back_populates="order")
    invoice = relationship("Invoice", back_populates="orders")

class Invoice(Base):
    __tablename__ = "invoices"
    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(Integer, ForeignKey("customers.id"), nullable=True)
    invoice_number = Column(String, unique=True)
    issue_date = Column(DateTime, default=datetime.datetime.utcnow)
    due_date = Column(DateTime)
    total_amount = Column(Float)
    discount_rate = Column(Float, default=0.0)
    is_bulk_discount = Column(Boolean, default=False)
    status = Column(Enum(InvoiceStatus), default=InvoiceStatus.UNPAID)
    delivery_status = Column(String, default="UNSENT") # UNSENT, SENT, MAILED
    memo = Column(String)

    orders = relationship("Order", back_populates="invoice")
    customer = relationship("Customer")

class AgencyOrder(Base):
    """代理店からの発注"""
    __tablename__ = "agency_orders"
    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(Integer, ForeignKey("customers.id"))
    order_number = Column(String, unique=True)
    order_date = Column(DateTime, default=datetime.datetime.utcnow)
    total_amount = Column(Float, default=0.0)
    status = Column(String, default="未処理")  # 未処理 / 処理済み / キャンセル
    invoice_id = Column(Integer, ForeignKey("invoices.id", ondelete="SET NULL"), nullable=True)
    memo = Column(String, nullable=True)

    customer = relationship("Customer", back_populates="agency_orders")
    items = relationship("AgencyOrderItem", back_populates="agency_order", cascade="all, delete-orphan")
    invoice = relationship("Invoice", backref="agency_orders")

class AgencyOrderItem(Base):
    """代理店発注の明細"""
    __tablename__ = "agency_order_items"
    id = Column(Integer, primary_key=True, index=True)
    agency_order_id = Column(Integer, ForeignKey("agency_orders.id", ondelete="CASCADE"))
    product_id = Column(Integer, ForeignKey("products.id", ondelete="SET NULL"), nullable=True)
    product_name = Column(String)
    quantity = Column(Integer)
    unit_price = Column(Float)
    subtotal = Column(Float)

    agency_order = relationship("AgencyOrder", back_populates="items")
    product = relationship("Product")

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

class Location(Base):
    """拠点（保管先）"""
    __tablename__ = "locations"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True, unique=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    stock_levels = relationship("ProductLocationStock", back_populates="location")

class ProductLocationStock(Base):
    """商品ごとの拠点別在庫"""
    __tablename__ = "product_location_stocks"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id", ondelete="CASCADE"))
    location_id = Column(Integer, ForeignKey("locations.id", ondelete="CASCADE"))
    stock_quantity = Column(Integer, default=0)

    product = relationship("Product", back_populates="location_stocks")
    location = relationship("Location", back_populates="stock_levels")

class StockMovement(Base):
    """在庫移動履歴"""
    __tablename__ = "stock_movements"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id", ondelete="CASCADE"))
    from_location_id = Column(Integer, ForeignKey("locations.id", ondelete="SET NULL"), nullable=True)
    to_location_id = Column(Integer, ForeignKey("locations.id", ondelete="SET NULL"), nullable=True)
    quantity = Column(Integer)
    type = Column(String) # "INBOUND", "OUTBOUND", "TRANSFER", "ADJUSTMENT"
    reason = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    product = relationship("Product")
    from_location = relationship("Location", foreign_keys=[from_location_id])
    to_location = relationship("Location", foreign_keys=[to_location_id])

# 既存の Product と関係性を紐づけるために Product クラスを更新（Relationshipを追加）
# (注: Productクラスの定義位置に戻り、relationshipを追加する必要があります)
