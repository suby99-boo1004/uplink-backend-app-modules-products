from typing import Optional, List, Dict, Any
from decimal import Decimal
from app.extensions import db
from .model import Product, ProductCategory, ProductItem
from .excel import read_products_from_excel, write_products_to_excel

def auto_calc_prices(price_design: int, price_small: Optional[int], price_delivery: Optional[int]) -> Dict[str, int]:
    """자동계산 규칙:
    - 소보수가 = 설계가 * 0.85 (단, 사용자가 입력하면 그 값을 우선)
    - 납품가 = 소보수가 * 0.82 (단, 사용자가 입력하면 그 값을 우선)
    """
    design = int(price_design or 0)
    small = int(price_small) if price_small is not None else int(round(design * 0.85))
    delivery = int(price_delivery) if price_delivery is not None else int(round(small * 0.82))
    return {"price_small": small, "price_delivery": delivery}

def get_or_create_category(name: str) -> Optional[ProductCategory]:
    n = (name or "").strip()
    if not n:
        return None
    cat = ProductCategory.query.filter_by(name=n).first()
    if cat:
        return cat
    cat = ProductCategory(name=n, is_active=True)
    db.session.add(cat)
    db.session.flush()
    return cat

def get_or_create_item(category_id: int, name: str) -> Optional[ProductItem]:
    n = (name or "").strip()
    if not category_id or not n:
        return None
    it = ProductItem.query.filter_by(category_id=category_id, name=n).first()
    if it:
        return it
    it = ProductItem(category_id=category_id, name=n, is_active=True)
    db.session.add(it)
    db.session.flush()
    return it

def list_products(q: str = "", category_id: Optional[int]=None, item_id: Optional[int]=None) -> List[Dict[str, Any]]:
    query = Product.query.filter_by(is_active=True)
    if q:
        query = query.filter(Product.name.ilike(f"%{q}%"))
    if category_id:
        query = query.filter(Product.category_id == category_id)
    if item_id:
        query = query.filter(Product.item_id == item_id)
    rows = query.order_by(Product.id.desc()).limit(2000).all()
    out = []
    for p in rows:
        out.append({
            "id": p.id,
            "category_id": p.category_id,
            "item_id": p.item_id,
            "name": p.name,
            "spec": p.spec,
            "price_design": int(p.price_design or 0),
            "price_small": int(p.price_small or 0),
            "price_delivery": int(p.price_delivery or 0),
            "stock_qty": int(p.stock_qty or 0),
            "memo": p.memo or ""
        })
    return out

def upsert_product(payload: Dict[str, Any], updated_by: Optional[int]=None, product_id: Optional[int]=None) -> Product:
    # 카테고리/항목은 이름으로도 들어올 수 있게(엑셀 업로드 대비)
    category_id = payload.get("category_id")
    item_id = payload.get("item_id")

    category_name = payload.get("category_name")
    item_name = payload.get("item_name")

    cat = None
    if not category_id and category_name:
        cat = get_or_create_category(category_name)
        category_id = cat.id if cat else None
    if category_id and item_name and not item_id:
        it = get_or_create_item(category_id, item_name)
        item_id = it.id if it else None

    price_design = int(payload.get("price_design") or 0)
    # 수동 수정 가능: 값이 들어오면 그 값을 우선, 없으면 자동
    ps_in = payload.get("price_small")
    pd_in = payload.get("price_delivery")
    auto = auto_calc_prices(price_design, ps_in if ps_in is not None else None, pd_in if pd_in is not None else None)

    if product_id:
        p = Product.query.get(product_id)
        if not p:
            raise ValueError("제품을 찾을 수 없습니다.")
    else:
        p = Product()

    p.category_id = category_id
    p.item_id = item_id
    p.name = (payload.get("name") or "").strip()
    if not p.name:
        raise ValueError("제품명(모델명)은 필수입니다.")
    p.spec = (payload.get("spec") or "").strip() or None
    p.price_design = price_design
    p.price_small = int(ps_in) if ps_in is not None else auto["price_small"]
    p.price_delivery = int(pd_in) if pd_in is not None else auto["price_delivery"]
    p.stock_qty = int(payload.get("stock_qty") or 0)
    p.memo = (payload.get("memo") or "").strip() or None
    p.updated_by = updated_by

    db.session.add(p)
    db.session.commit()
    return p

def soft_delete_product(product_id: int):
    p = Product.query.get(product_id)
    if not p:
        raise ValueError("제품을 찾을 수 없습니다.")
    p.is_active = False
    db.session.commit()

def upload_excel(file_bytes: bytes, updated_by: Optional[int]=None) -> Dict[str, Any]:
    rows, errors = read_products_from_excel(file_bytes)
    if errors:
        return {"ok": False, "errors": errors, "imported": 0}

    imported = 0
    for r in rows:
        # 엑셀은 소보수가/납품가를 이미 갖고 있을 수 있음(수동값 우선)
        try:
            upsert_product(r, updated_by=updated_by, product_id=None)
            imported += 1
        except Exception as e:
            errors.append(str(e))

    return {"ok": len(errors) == 0, "errors": errors, "imported": imported}

def download_excel() -> bytes:
    # 엑셀 다운로드는 템플릿(항목/구분/모델명...) 형태로 내보냄
    products = list_products()
    # category/item 이름도 채워 넣기 위해 join
    cat_map = {c.id: c.name for c in ProductCategory.query.all()}
    item_map = {(i.id): i for i in ProductItem.query.all()}

    rows = []
    for p in products:
        rows.append({
            "item_name": item_map.get(p["item_id"]).name if p["item_id"] in item_map else "",
            "category_name": cat_map.get(p["category_id"], ""),
            "name": p["name"],
            "spec": p["spec"],
            "price_design": p["price_design"],
            "price_small": p["price_small"],
            "price_delivery": p["price_delivery"],
            "stock_qty": p["stock_qty"],
            "memo": p["memo"],
        })
    return write_products_to_excel(rows)
