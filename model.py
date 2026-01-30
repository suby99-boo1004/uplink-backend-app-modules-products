from app.extensions import db

class ProductCategory(db.Model):
    __tablename__ = "product_categories"
    id = db.Column(db.BigInteger, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    is_active = db.Column(db.Boolean, nullable=False, default=True)

class ProductItem(db.Model):
    __tablename__ = "product_items"
    id = db.Column(db.BigInteger, primary_key=True)
    category_id = db.Column(db.BigInteger, db.ForeignKey("product_categories.id"), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    __table_args__ = (db.UniqueConstraint("category_id", "name", name="uq_product_items_category_name"),)

class Product(db.Model):
    __tablename__ = "products"
    id = db.Column(db.BigInteger, primary_key=True)
    category_id = db.Column(db.BigInteger, db.ForeignKey("product_categories.id"))
    item_id = db.Column(db.BigInteger, db.ForeignKey("product_items.id"))

    # 템플릿 기준: 모델명 = 제품명
    name = db.Column(db.String(200), nullable=False)
    spec = db.Column(db.String(200))

    price_design = db.Column(db.Numeric(14, 0), nullable=False, default=0)
    price_small = db.Column(db.Numeric(14, 0), nullable=False, default=0)     # 설계가*0.85 (자동)
    price_delivery = db.Column(db.Numeric(14, 0), nullable=False, default=0)  # 소보수가*0.82 (자동)

    stock_qty = db.Column(db.Integer, nullable=False, default=0)
    memo = db.Column(db.Text)
    is_active = db.Column(db.Boolean, nullable=False, default=True)

    updated_by = db.Column(db.BigInteger)  # users.id (선택: FK로 걸어도 됨)
    updated_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), onupdate=db.func.now())

class ProductStockLog(db.Model):
    __tablename__ = "product_stock_logs"
    id = db.Column(db.BigInteger, primary_key=True)
    product_id = db.Column(db.BigInteger, db.ForeignKey("products.id"), nullable=False)
    delta_qty = db.Column(db.Integer, nullable=False)
    reason = db.Column(db.String(40), nullable=False)  # ENUM을 이미 사용 중이면 변경
    ref_table = db.Column(db.String(60))
    ref_id = db.Column(db.BigInteger)
    before_qty = db.Column(db.Integer, nullable=False)
    after_qty = db.Column(db.Integer, nullable=False)
    created_by = db.Column(db.BigInteger)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now())
