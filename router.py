from __future__ import annotations

from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from fastapi.responses import StreamingResponse
from typing import Optional, List, Dict, Any, Tuple
import os
import json
import re
import io
import datetime

import openpyxl

# =========================================================
# 업링크 제품(자재관리) Router (파일 저장형 - v1.13)
#
# ✅ 권한(role_id) 고정 매핑 (대표님 확정)
# - 관리자   ADMIN    = 6
# - 운영자   OPERATOR = 7
# - 회사직원 STAFF    = 8
# - 외부직원 EXTERNAL = 9
# - 게스트   GUEST    = 10
#
# 정책:
# - 조회(GET /api/products): 로그인 없이도 가능(현재 구조 유지)
# - 등록/수정/삭제/업로드/다운로드: 관리자(6) 또는 운영자(7)만 허용
#
# 전제(routes.py):
#   app.include_router(products_router, prefix="/api/products", tags=["products"])
# =========================================================

router = APIRouter(tags=["products"])

from app.core.deps import get_current_user
from app.models.user import User

ROLE_ADMIN_ID = 6
ROLE_OPERATOR_ID = 7
ROLE_STAFF_ID = 8
ROLE_EXTERNAL_ID = 9
ROLE_GUEST_ID = 10

DATA_DIR = os.path.join(os.path.dirname(__file__), "_data")
DATA_FILE = os.path.join(DATA_DIR, "products.json")

EXPECTED_HEADERS = ["항목", "구분", "모델명", "규격", "설계가", "수의계약가", "납품가", "사무실 재고", "비고"]


def _role_id(user: Any) -> Optional[int]:
    rid = getattr(user, "role_id", None)
    try:
        return int(rid) if rid is not None else None
    except Exception:
        return None


def require_admin_or_operator(user: User = Depends(get_current_user)) -> User:
    rid = _role_id(user)
    if rid in (ROLE_ADMIN_ID, ROLE_OPERATOR_ID):
        return user
    raise HTTPException(status_code=403, detail=f"관리자/운영자만 가능합니다. (role_id={rid})")


def _atomic_write(store: Dict[str, Any]):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp_path = DATA_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, DATA_FILE)


def _ensure_store():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DATA_FILE):
        _atomic_write({"seq": 0, "rows": []})


def _load_store() -> Dict[str, Any]:
    _ensure_store()
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_store(store: Dict[str, Any]):
    _ensure_store()
    _atomic_write(store)


def _parse_money(v) -> int:
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    s = re.sub(r"[^0-9]", "", s)
    return int(s) if s else 0


def _parse_int(v) -> int:
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    s = re.sub(r"[^0-9-]", "", s)
    return int(s) if s else 0


def _cut_to_thousand(value: int) -> int:
    # 백단위 절사 → 천 단위부터 표시
    return (value // 1000) * 1000


def _auto_prices(price_design: int, price_small: Optional[int], price_delivery: Optional[int]) -> Dict[str, int]:
    # 소보수가 = 설계가 * 0.85 (천 단위 절사)
    if price_small is None:
        raw_small = int(price_design * 0.85)
        small = _cut_to_thousand(raw_small)
    else:
        small = _cut_to_thousand(int(price_small))

    # 납품가(수의계약가) = 소보수가 * 0.83 (천 단위 절사)
    if price_delivery is None:
        raw_delivery = int(small * 0.83)
        delivery = _cut_to_thousand(raw_delivery)
    else:
        delivery = _cut_to_thousand(int(price_delivery))

    return {"price_small": small, "price_delivery": delivery}


def _sort_key(r: Dict[str, Any]) -> Tuple[str, str, str, str]:
    # 항목 -> 구분 -> 모델명 -> 규격
    return (
        (r.get("item_name") or "").strip(),
        (r.get("category_name") or "").strip(),
        (r.get("name") or "").strip(),
        (r.get("spec") or "").strip(),
    )


def _match_q(row: Dict[str, Any], q: str) -> bool:
    qq = q.lower()
    # 검색 대상: 항목/구분/모델명/규격/비고
    for k in ("item_name", "category_name", "name", "spec", "memo"):
        v = row.get(k)
        if v and qq in str(v).lower():
            return True
    return False


def _list_products(q: Optional[str]) -> List[Dict[str, Any]]:
    store = _load_store()
    rows = store.get("rows", [])
    if q and str(q).strip():
        rows = [r for r in rows if _match_q(r, str(q))]
    return sorted(rows, key=_sort_key)


def _find_by_id(rows: List[Dict[str, Any]], product_id: int) -> Optional[Dict[str, Any]]:
    for r in rows:
        if int(r.get("id", 0)) == int(product_id):
            return r
    return None


# ---------- READ ----------
@router.get("")
def list_products_no_slash(q: Optional[str] = None) -> List[Dict[str, Any]]:
    return _list_products(q)


@router.get("/")
def list_products(q: Optional[str] = None) -> List[Dict[str, Any]]:
    return _list_products(q)


# ---------- WRITE (ADMIN/OPERATOR only) ----------
@router.post("", dependencies=[Depends(require_admin_or_operator)])
def create_product(payload: Dict[str, Any]) -> Dict[str, Any]:
    name = (payload.get("name") or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="모델명(name)은 필수입니다.")

    item_name = (payload.get("item_name") or "").strip()
    category_name = (payload.get("category_name") or "").strip()
    spec = (payload.get("spec") or "").strip()
    memo = (payload.get("memo") or "").strip()

    price_design = int(payload.get("price_design") or 0)
    ps_in = payload.get("price_small")
    pd_in = payload.get("price_delivery")
    auto = _auto_prices(price_design, int(ps_in) if ps_in is not None else None, int(pd_in) if pd_in is not None else None)

    store = _load_store()
    rows = store.get("rows", [])
    seq = int(store.get("seq", 0)) + 1

    row = {
        "id": seq,
        "item_name": item_name,
        "category_name": category_name,
        "name": name,
        "spec": spec,
        "price_design": price_design,
        "price_small": auto["price_small"],
        "price_delivery": auto["price_delivery"],
        "stock_qty": int(payload.get("stock_qty") or 0),
        "memo": memo,
    }

    rows.append(row)
    store["seq"] = seq
    store["rows"] = rows
    _save_store(store)
    return {"ok": True, "row": row}


@router.patch("/{product_id}", dependencies=[Depends(require_admin_or_operator)])
def update_product(product_id: int, payload: Dict[str, Any]) -> Dict[str, Any]:
    store = _load_store()
    rows = store.get("rows", [])
    target = _find_by_id(rows, product_id)
    if not target:
        raise HTTPException(status_code=404, detail="제품을 찾을 수 없습니다.")

    for k in ("item_name", "category_name", "name", "spec", "memo"):
        if k in payload and payload[k] is not None:
            target[k] = str(payload[k]).strip()

    if "stock_qty" in payload and payload["stock_qty"] is not None:
        target["stock_qty"] = int(payload["stock_qty"])

    if any(k in payload for k in ("price_design", "price_small", "price_delivery")):
        price_design = int(payload.get("price_design", target.get("price_design", 0)) or 0)
        ps_in = payload.get("price_small")
        pd_in = payload.get("price_delivery")
        auto = _auto_prices(price_design, int(ps_in) if ps_in is not None else None, int(pd_in) if pd_in is not None else None)

        target["price_design"] = price_design
        target["price_small"] = auto["price_small"]
        target["price_delivery"] = auto["price_delivery"]

    _save_store(store)
    return {"ok": True, "row": target}


@router.delete("/{product_id}", dependencies=[Depends(require_admin_or_operator)])
def delete_product(product_id: int) -> Dict[str, Any]:
    store = _load_store()
    rows = store.get("rows", [])
    before = len(rows)
    rows = [r for r in rows if int(r.get("id", 0)) != int(product_id)]
    if len(rows) == before:
        raise HTTPException(status_code=404, detail="제품을 찾을 수 없습니다.")
    store["rows"] = rows
    _save_store(store)
    return {"ok": True}


@router.post("/upload", dependencies=[Depends(require_admin_or_operator)])
async def upload_products_excel(file: UploadFile = File(...)) -> Dict[str, Any]:
    content = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")

    ws = wb.active
    headers = [str(ws.cell(1, c).value).strip() if ws.cell(1, c).value is not None else "" for c in range(1, 10)]
    if headers[:9] != EXPECTED_HEADERS:
        raise HTTPException(status_code=400, detail=f"엑셀 헤더가 템플릿과 다릅니다. 기대={EXPECTED_HEADERS}, 실제={headers[:9]}")

    incoming: List[Dict[str, Any]] = []
    errors: List[str] = []

    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, 10)]
        if all(v is None or str(v).strip() == "" for v in row):
            continue

        item_name = str(row[0]).strip() if row[0] is not None else ""
        category_name = str(row[1]).strip() if row[1] is not None else ""
        name = str(row[2]).strip() if row[2] is not None else ""
        spec = str(row[3]).strip() if row[3] is not None else ""

        if not name:
            errors.append(f"{r}행: 모델명(제품명)이 비어있습니다.")
            continue

        price_design = int(_parse_money(row[4]))
        price_small_raw = int(_parse_money(row[5]))
        price_delivery_raw = int(_parse_money(row[6]))

        ps_in = price_small_raw if price_small_raw > 0 else None
        pd_in = price_delivery_raw if price_delivery_raw > 0 else None
        auto = _auto_prices(price_design, ps_in, pd_in)

        stock_qty = _parse_int(row[7])
        memo = str(row[8]).strip() if row[8] is not None else ""

        incoming.append(
            {
                "item_name": item_name,
                "category_name": category_name,
                "name": name,
                "spec": spec,
                "price_design": price_design,
                "price_small": auto["price_small"],
                "price_delivery": auto["price_delivery"],
                "stock_qty": stock_qty,
                "memo": memo,
            }
        )

    if errors:
        return {"ok": False, "imported": 0, "errors": errors}

    store = _load_store()
    rows = store.get("rows", [])
    seq = int(store.get("seq", 0))

    # upsert key = (모델명, 규격)
    by_key = {(str(r.get("name") or ""), str(r.get("spec") or "")): r for r in rows}

    imported = 0
    for inc in incoming:
        k = (str(inc.get("name") or ""), str(inc.get("spec") or ""))
        existing = by_key.get(k)
        if existing:
            existing.update(inc)
        else:
            seq += 1
            rows.append({"id": seq, **inc})
        imported += 1

    store["seq"] = seq
    store["rows"] = rows
    _save_store(store)
    return {"ok": True, "imported": imported, "errors": []}


@router.get("/download", dependencies=[Depends(require_admin_or_operator)])
def download_products_excel() -> StreamingResponse:
    store = _load_store()
    rows = sorted(store.get("rows", []), key=_sort_key)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for c, h in enumerate(EXPECTED_HEADERS, start=1):
        ws.cell(1, c).value = h

    for i, r in enumerate(rows, start=2):
        ws.cell(i, 1).value = r.get("item_name", "")
        ws.cell(i, 2).value = r.get("category_name", "")
        ws.cell(i, 3).value = r.get("name", "")
        ws.cell(i, 4).value = r.get("spec", "")
        ws.cell(i, 5).value = int(r.get("price_design", 0) or 0)
        ws.cell(i, 6).value = int(r.get("price_small", 0) or 0)
        ws.cell(i, 7).value = int(r.get("price_delivery", 0) or 0)
        ws.cell(i, 8).value = int(r.get("stock_qty", 0) or 0)
        ws.cell(i, 9).value = r.get("memo", "")

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"products_export_{datetime.date.today().isoformat()}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
