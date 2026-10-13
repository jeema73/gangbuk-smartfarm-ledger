import os
import json
from datetime import datetime
from typing import List, Dict, Any, Optional
import firebase_admin
from firebase_admin import credentials, firestore

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FIREBASE_KEY_PATH = os.path.join(BASE_DIR, "firebase_key.json")

_db = None

def get_db():
    """Firestore 클라이언트 싱글톤 반환 (로컬 파일 또는 환경변수 서비스 키 지원)"""
    global _db
    if _db is not None:
        return _db

    if not firebase_admin._apps:
        # 1. 파일 시스템에서 firebase_key.json 확인
        if os.path.exists(FIREBASE_KEY_PATH):
            cred = credentials.Certificate(FIREBASE_KEY_PATH)
            firebase_admin.initialize_app(cred)
        # 2. Render 환경변수 FIREBASE_KEY_JSON 지원
        elif os.environ.get("FIREBASE_KEY_JSON"):
            key_dict = json.loads(os.environ["FIREBASE_KEY_JSON"])
            cred = credentials.Certificate(key_dict)
            firebase_admin.initialize_app(cred)
        else:
            raise RuntimeError("firebase_key.json 파일을 찾을 수 없습니다.")

    _db = firestore.client()
    return _db

# ========================================================
# 1. 공급자 및 담당자/서명 정보 (Settings)
# ========================================================
DEFAULT_SUPPLIER = {
    "name": "강북구청",
    "department": "지역경제과",
    "biz_no": "210-83-02118",
    "address": "서울 강북구 덕릉로 138 (번동) 강북구청",
    "manager": "권준형 주무관",
    "tel": "02-901-7256",
    "delivery_manager": "박기남",
    "delivery_tel": "010-0000-0000",
    "signature_image": ""  # Base64 data URL
}

def get_supplier_info() -> Dict[str, Any]:
    try:
        db = get_db()
        doc = db.collection("settings").document("supplier").get()
        if doc.exists:
            data = doc.to_dict()
            res = DEFAULT_SUPPLIER.copy()
            res.update(data)
            return res
    except Exception as e:
        print("[Firestore get_supplier_info Error]:", e)
    return DEFAULT_SUPPLIER.copy()

def save_supplier_info(info: Dict[str, Any]) -> Dict[str, Any]:
    db = get_db()
    current = get_supplier_info()
    current.update(info)
    current["updated_at"] = datetime.now().isoformat()
    db.collection("settings").document("supplier").set(current)
    return {"success": True, "message": "공급자 및 담당자 정보가 저장되었습니다."}

# ========================================================
# 2. 작물 품목 및 기본 단가 (Crops)
# ========================================================
DEFAULT_CROPS = [
    {"name": "루꼴라", "unit_price": 17.7, "category": "허브류", "unit": "g", "remarks": "생잎/피자용"},
    {"name": "엽채류", "unit_price": 11.5, "category": "쌈채소", "unit": "g", "remarks": "샐러드용"},
    {"name": "애플민트", "unit_price": 56.6, "category": "허브류", "unit": "g", "remarks": "음료/모히또"},
    {"name": "바질", "unit_price": 35.0, "category": "허브류", "unit": "g", "remarks": "스위트바질"},
    {"name": "딸기(설향)", "unit_price": 10.0, "category": "과채류", "unit": "g", "remarks": "생과"},
    {"name": "애플수박", "unit_price": 8.0, "category": "과채류", "unit": "g", "remarks": "생과"}
]

def init_default_crops_if_empty():
    db = get_db()
    crops_ref = db.collection("crops")
    docs = list(crops_ref.limit(1).stream())
    if not docs:
        for c in DEFAULT_CROPS:
            crops_ref.document(c["name"]).set(c)

def get_all_crops() -> List[Dict[str, Any]]:
    try:
        db = get_db()
        init_default_crops_if_empty()
        docs = db.collection("crops").stream()
        crops = [d.to_dict() for d in docs]
        return sorted(crops, key=lambda x: x.get("name", ""))
    except Exception as e:
        print("[Firestore get_all_crops Error]:", e)
        return DEFAULT_CROPS.copy()

def add_or_update_crop(crop_data: Dict[str, Any]) -> Dict[str, Any]:
    db = get_db()
    name = crop_data.get("name", "").strip()
    if not name:
        raise ValueError("작물명을 입력해주세요.")
    unit_price = float(crop_data.get("unit_price", 0))
    original_name = crop_data.get("original_name", "").strip()

    if original_name and original_name != name:
        db.collection("crops").document(original_name).delete()

    data = {
        "name": name,
        "unit_price": unit_price,
        "category": crop_data.get("category", "채소류"),
        "unit": crop_data.get("unit", "g"),
        "remarks": crop_data.get("remarks", "").strip(),
        "updated_at": datetime.now().isoformat()
    }
    db.collection("crops").document(name).set(data)
    return {"success": True, "message": f"'{name}' 작물이 성공적으로 저장되었습니다."}

def delete_crop(crop_name: str) -> Dict[str, Any]:
    db = get_db()
    db.collection("crops").document(crop_name.strip()).delete()
    return {"success": True, "message": f"'{crop_name}' 작물이 삭제되었습니다."}

# ========================================================
# 3. 협력 매장 목록 (Stores)
# ========================================================
DEFAULT_STORES = [
    {"name": "419국수김밥", "address": "강북구 4.19로 31", "menu": "루꼴라 크림치즈김밥", "price": "5000"},
    {"name": "미즐", "address": "강북구 4.19로 16", "menu": "루꼴라 돈까스", "price": "9000"},
    {"name": "까미노", "address": "강북구 4.19로 61", "menu": "루꼴라 떡볶이", "price": "5500"},
    {"name": "비르케", "address": "강북구 4.19로 61", "menu": "루꼴라 샐러드", "price": "9000"},
    {"name": "화수브루어리", "address": "강북구 4.19로 45", "menu": "루꼴라 피자", "price": "18000"},
    {"name": "미즐 카페엠", "address": "강북구 4.19로 16", "menu": "루꼴라 샐러드", "price": "12000"},
    {"name": "민지네", "address": "강북구 4.19로 65", "menu": "루꼴라 전", "price": "8000"}
]

def init_default_stores_if_empty():
    db = get_db()
    stores_ref = db.collection("stores")
    docs = list(stores_ref.limit(1).stream())
    if not docs:
        for s in DEFAULT_STORES:
            stores_ref.document(s["name"]).set(s)

def get_all_stores() -> List[Dict[str, Any]]:
    try:
        db = get_db()
        init_default_stores_if_empty()
        docs = db.collection("stores").stream()
        stores = [d.to_dict() for d in docs]
        return sorted(stores, key=lambda x: x.get("name", ""))
    except Exception as e:
        print("[Firestore get_all_stores Error]:", e)
        return DEFAULT_STORES.copy()

def add_or_update_store(store_data: Dict[str, Any]) -> Dict[str, Any]:
    db = get_db()
    name = store_data.get("name", "").strip()
    if not name:
        raise ValueError("매장명을 입력해주세요.")
    original_name = store_data.get("original_name", "").strip()

    if original_name and original_name != name:
        db.collection("stores").document(original_name).delete()

    data = {
        "name": name,
        "address": store_data.get("address", "").strip(),
        "menu": store_data.get("menu", "").strip(),
        "price": str(store_data.get("price", "")).strip(),
        "manager": store_data.get("manager", "").strip(),
        "tel": store_data.get("tel", "").strip(),
        "updated_at": datetime.now().isoformat()
    }
    db.collection("stores").document(name).set(data)
    return {"success": True, "message": f"'{name}' 매장이 성공적으로 저장되었습니다."}

def delete_store(store_name: str) -> Dict[str, Any]:
    db = get_db()
    db.collection("stores").document(store_name.strip()).delete()
    return {"success": True, "message": f"'{store_name}' 매장이 삭제되었습니다."}

def get_store(store_name: str) -> Optional[Dict[str, Any]]:
    db = get_db()
    doc = db.collection("stores").document(store_name.strip()).get()
    return doc.to_dict() if doc.exists else None

# ========================================================
# 4. 납품 주문 내역 (Orders)
# ========================================================
def add_delivery_order(data: Dict[str, Any]) -> Dict[str, Any]:
    db = get_db()
    order_date = data.get("order_date", "").strip()
    delivery_date = data.get("delivery_date") or order_date
    delivery_date = delivery_date.strip()
    store_name = data.get("store_name", "").strip()
    items = data.get("items", [])

    if not order_date:
        raise ValueError("주문일자를 입력해주세요.")
    if not store_name:
        raise ValueError("매장명을 선택해주세요.")
    if not items:
        raise ValueError("최소 1개 이상의 품목을 입력해주세요.")

    total_qty = 0.0
    total_amount = 0.0
    cleaned_items = []
    for it in items:
        q = float(it.get("qty", 0))
        p = float(it.get("unit_price", 0))
        amt = round(q * p)
        total_qty += q
        total_amount += amt
        cleaned_items.append({
            "crop": it.get("crop", "").strip(),
            "qty": q,
            "unit_price": p,
            "amount": amt,
            "remarks": it.get("remarks", "").strip()
        })

    order_key = f"{order_date}___{store_name}"
    doc_id = f"{delivery_date}_{store_name}_{datetime.now().strftime('%H%M%S')}"

    order_doc = {
        "id": doc_id,
        "order_key": order_key,
        "order_date": order_date,
        "delivery_date": delivery_date,
        "store_name": store_name,
        "items": cleaned_items,
        "total_qty": total_qty,
        "total_amount": total_amount,
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat()
    }
    db.collection("orders").document(doc_id).set(order_doc)

    return {
        "success": True,
        "order_key": order_key,
        "order_id": doc_id,
        "message": f"'{store_name}' ({len(items)}개 품목) 납품 등록이 완료되었습니다."
    }

def update_delivery_order(target_key: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """order_key(또는 doc_id)로 기존 주문을 찾아 수정"""
    db = get_db()
    new_order_date = data.get("order_date", "").strip()
    new_delivery_date = data.get("delivery_date") or new_order_date
    new_delivery_date = new_delivery_date.strip()
    new_store_name = data.get("store_name", "").strip()
    items = data.get("items", [])

    if not items:
        raise ValueError("최소 1개 이상의 품목을 입력해주세요.")

    # 1. 일치하는 기존 문서 탐색
    orders_ref = db.collection("orders")
    matched_docs = []

    # id 직접 검색
    doc_direct = orders_ref.document(target_key).get()
    if doc_direct.exists:
        matched_docs.append(doc_direct)
    else:
        # order_key로 검색
        query = orders_ref.where("order_key", "==", target_key).stream()
        matched_docs = list(query)

    if not matched_docs:
        # 날짜_매장명 분리 검색
        parts = target_key.split("___")
        if len(parts) == 2:
            t_date, t_store = parts[0].strip(), parts[1].strip()
            q2 = orders_ref.where("store_name", "==", t_store).where("order_date", "==", t_date).stream()
            matched_docs = list(q2)
            if not matched_docs:
                q3 = orders_ref.where("store_name", "==", t_store).where("delivery_date", "==", t_date).stream()
                matched_docs = list(q3)

    if not matched_docs:
        raise ValueError(f"수정할 주문 내역('{target_key}')을 찾을 수 없습니다.")

    target_doc_id = matched_docs[0].id

    total_qty = 0.0
    total_amount = 0.0
    cleaned_items = []
    for it in items:
        q = float(it.get("qty", 0))
        p = float(it.get("unit_price", 0))
        amt = round(q * p)
        total_qty += q
        total_amount += amt
        cleaned_items.append({
            "crop": it.get("crop", "").strip(),
            "qty": q,
            "unit_price": p,
            "amount": amt,
            "remarks": it.get("remarks", "").strip()
        })

    new_order_key = f"{new_order_date}___{new_store_name}"
    update_data = {
        "order_key": new_order_key,
        "order_date": new_order_date,
        "delivery_date": new_delivery_date,
        "store_name": new_store_name,
        "items": cleaned_items,
        "total_qty": total_qty,
        "total_amount": total_amount,
        "updated_at": datetime.now().isoformat()
    }
    orders_ref.document(target_doc_id).update(update_data)

    return {
        "success": True,
        "order_key": new_order_key,
        "order_id": target_doc_id,
        "message": f"'{new_store_name}' 납품 내역이 성공적으로 수정되었습니다."
    }

def delete_delivery_order(target_key: str) -> Dict[str, Any]:
    db = get_db()
    orders_ref = db.collection("orders")
    matched_docs = []

    doc_direct = orders_ref.document(target_key).get()
    if doc_direct.exists:
        matched_docs.append(doc_direct)
    else:
        query = orders_ref.where("order_key", "==", target_key).stream()
        matched_docs = list(query)
        if not matched_docs:
            parts = target_key.split("___")
            if len(parts) == 2:
                t_date, t_store = parts[0].strip(), parts[1].strip()
                q2 = orders_ref.where("store_name", "==", t_store).where("order_date", "==", t_date).stream()
                matched_docs = list(q2)
                if not matched_docs:
                    q3 = orders_ref.where("store_name", "==", t_store).where("delivery_date", "==", t_date).stream()
                    matched_docs = list(q3)

    if not matched_docs:
        raise ValueError(f"삭제할 주문 내역('{target_key}')을 찾을 수 없습니다.")

    for d in matched_docs:
        orders_ref.document(d.id).delete()

    return {"success": True, "message": "주문 내역이 성공적으로 삭제되었습니다."}

def get_orders(search_date: Optional[str] = None, search_store: Optional[str] = None, year_month: Optional[str] = None, limit: int = 500) -> List[Dict[str, Any]]:
    """모든 통계와 정렬의 기본 기준은 delivery_date (납품일)"""
    try:
        db = get_db()
        orders_ref = db.collection("orders")
        docs = orders_ref.stream()
        orders = []
        for d in docs:
            obj = d.to_dict()
            orders.append(obj)

        filtered = []
        for o in orders:
            eff_date = o.get("delivery_date") or o.get("order_date", "")
            o_date = o.get("order_date", "")
            s_name = o.get("store_name", "")

            if year_month and not eff_date.startswith(year_month) and not o_date.startswith(year_month):
                continue
            if search_date and search_date != eff_date and search_date != o_date:
                continue
            if search_store and search_store not in s_name:
                continue
            filtered.append(o)

        # 납품일 최신순 내림차순 정렬
        sorted_orders = sorted(filtered, key=lambda x: (x.get("delivery_date") or "", x.get("created_at") or ""), reverse=True)
        return sorted_orders[:limit]
    except Exception as e:
        print("[Firestore get_orders Error]:", e)
        return []

def get_receipt_data(order_key: str) -> Dict[str, Any]:
    """확인서 발행용 상세 데이터 구성 (공급자, 납품자, 서명이미지, 한글금액 등 포함)"""
    db = get_db()
    orders_ref = db.collection("orders")
    matched_docs = []

    doc_direct = orders_ref.document(order_key).get()
    if doc_direct.exists:
        matched_docs.append(doc_direct)
    else:
        query = orders_ref.where("order_key", "==", order_key).stream()
        matched_docs = list(query)
        if not matched_docs:
            parts = order_key.split("___")
            if len(parts) == 2:
                t_date, t_store = parts[0].strip(), parts[1].strip()
                q2 = orders_ref.where("store_name", "==", t_store).where("order_date", "==", t_date).stream()
                matched_docs = list(q2)
                if not matched_docs:
                    q3 = orders_ref.where("store_name", "==", t_store).where("delivery_date", "==", t_date).stream()
                    matched_docs = list(q3)

    if not matched_docs:
        raise ValueError(f"확인서 발행용 주문 내역('{order_key}')을 찾을 수 없습니다.")

    order = matched_docs[0].to_dict()
    supplier = get_supplier_info()

    # 매장 주소 조회
    store_name = order.get("store_name", "")
    store_doc = db.collection("stores").document(store_name).get()
    address = store_doc.to_dict().get("address", "") if store_doc.exists else ""

    from delivery_manager import number_to_korean
    total_amt = int(round(order.get("total_amount", 0)))
    korean_amt = number_to_korean(total_amt)

    total_qty = order.get("total_qty", 0)
    formatted_total_qty = f"{total_qty:,.1f}".rstrip('0').rstrip('.') if total_qty else "0"

    return {
        "order_key": order.get("order_key", order_key),
        "order_date": order.get("order_date", ""),
        "delivery_date": order.get("delivery_date") or order.get("order_date", ""),
        "store_name": store_name,
        "address": address,
        "items": order.get("items", []),
        "total_qty": total_qty,
        "formatted_total_qty": formatted_total_qty,
        "total_amount": total_amt,
        "formatted_total_amount": f"{total_amt:,}",
        "korean_amount": korean_amt,
        "supplier": supplier
    }
