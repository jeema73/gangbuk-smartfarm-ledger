import openpyxl
import os
import shutil
import json
from datetime import datetime
from typing import List, Dict, Any, Optional

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "납품대장.xlsx")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

DEFAULT_SUPPLIER = {
    "name": "강북구청",
    "department": "스마트팜 사업단 / 도시농업",
    "biz_no": "109-83-00109",
    "address": "서울특별시 강북구 도봉로 89 (수유동, 강북구청)",
    "manager": "박기남",
    "tel": "02-901-6114"
}

def get_supplier_info() -> Dict[str, str]:
    # 1. 로컬 config.json 우선 로드
    supplier = DEFAULT_SUPPLIER.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                supplier.update(data.get("supplier", {}))
        except Exception:
            pass

    # 2. 구글 시트 납품대장 N:O열에 저장된 최신 공급자 정보가 있으면 자동 병합
    try:
        import google_sheets_manager
        sheet_supplier = google_sheets_manager.fetch_supplier_from_sheet()
        if sheet_supplier:
            supplier.update(sheet_supplier)
    except Exception:
        pass

    return supplier

def save_supplier_info(info: Dict[str, str]):
    # 1. config.json에 저장
    current_config = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                current_config = json.load(f)
        except Exception:
            pass
    current_config["supplier"] = info
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(current_config, f, ensure_ascii=False, indent=2)

    # 2. 로컬 엑셀 납품대장 시트 N:O 열에 공급자 정보 동기화
    try:
        ensure_excel_exists()
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
        if "납품대장" in wb.sheetnames:
            ws_d = wb["납품대장"]
            mapping = [
                ("상호/기관명", info.get("name", "")),
                ("부서", info.get("department", "")),
                ("사업자번호", info.get("biz_no", "")),
                ("소재지(주소)", info.get("address", "")),
                ("담당자", info.get("manager", "")),
                ("연락처", info.get("tel", ""))
            ]
            for idx, (lbl, val) in enumerate(mapping, start=1):
                ws_d.cell(idx, 14, lbl) # N열
                ws_d.cell(idx, 15, val) # O열
            wb.save(EXCEL_PATH)
    except Exception as e:
        print("[Excel Supplier Sync Warning]:", e)

    # 3. 구글 스프레드시트 Webhook 및 동기화 요청
    try:
        import google_sheets_manager
        conf = google_sheets_manager.get_google_sheet_config()
        webhook_url = conf.get("webhook_url")
        if webhook_url:
            import requests
            requests.post(webhook_url, json={"action": "save_supplier", "supplier": info}, timeout=10)
    except Exception as g_err:
        print("[Google Sheet Supplier Sync Warning]:", g_err)

DEFAULT_CROPS_LIST = [
    {"name": "루꼴라", "unit_price": 17.7, "category": "허브류", "unit": "g", "remarks": "생잎/피자용"},
    {"name": "엽채류", "unit_price": 11.5, "category": "쌈채소", "unit": "g", "remarks": "샐러드용"},
    {"name": "애플민트", "unit_price": 56.6, "category": "허브류", "unit": "g", "remarks": "음료/모히또"},
    {"name": "바질", "unit_price": 35.0, "category": "허브류", "unit": "g", "remarks": "스위트바질"},
    {"name": "딸기", "unit_price": 50.0, "category": "과채류", "unit": "g", "remarks": "생과"}
]

def get_crops_detailed() -> List[Dict[str, Any]]:
    """작물 품목 및 단가 목록 상세 조회 (config.json 및 납품대장 K, L열)"""
    crops_map = {c["name"]: dict(c) for c in DEFAULT_CROPS_LIST}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "crops" in data and isinstance(data["crops"], list):
                    stored_map = {}
                    for c in data["crops"]:
                        if isinstance(c, dict) and "name" in c:
                            stored_map[c["name"]] = c
                    if stored_map:
                        crops_map = stored_map
        except Exception:
            pass

    # 엑셀 K, L 열에서도 누락된 작물이 있으면 자동 반영
    try:
        if os.path.exists(EXCEL_PATH):
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
            if "납품대장" in wb.sheetnames:
                ws_d = wb["납품대장"]
                for r in range(2, 30):
                    c_name = ws_d.cell(r, 11).value
                    c_price = ws_d.cell(r, 12).value
                    if c_name:
                        name_str = str(c_name).strip()
                        if name_str not in crops_map:
                            try:
                                p = float(c_price) if c_price is not None else 0.0
                            except (ValueError, TypeError):
                                p = 0.0
                            crops_map[name_str] = {
                                "name": name_str,
                                "unit_price": p,
                                "category": "일반작물",
                                "unit": "g",
                                "remarks": ""
                            }
    except Exception:
        pass

    return list(crops_map.values())

def save_crops_list(crops_list: List[Dict[str, Any]]):
    """작물 목록을 config.json 및 엑셀 납품대장 시트(K, L열)에 동기화 저장"""
    current_config = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                current_config = json.load(f)
        except Exception:
            pass
    current_config["crops"] = crops_list
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(current_config, f, ensure_ascii=False, indent=2)

    # 엑셀 파일 동기화
    try:
        ensure_excel_exists()
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
        if "납품대장" in wb.sheetnames:
            ws_d = wb["납품대장"]
            ws_d.cell(1, 11).value = "작물명"
            ws_d.cell(1, 12).value = "단가(1g)"
            for r in range(2, 50):
                ws_d.cell(r, 11).value = None
                ws_d.cell(r, 12).value = None
            for idx, c in enumerate(crops_list):
                r = idx + 2
                ws_d.cell(r, 11).value = c["name"]
                ws_d.cell(r, 12).value = float(c.get("unit_price", 0.0))
            wb.save(EXCEL_PATH)
    except Exception as e:
        print("[Warning] 작물 엑셀 동기화 예외:", e)

    try:
        import threading
        import google_sheets_manager
        if google_sheets_manager.is_google_sheets_configured():
            threading.Thread(target=google_sheets_manager.sync_full_ledger_to_sheet, daemon=True).start()
    except Exception:
        pass

def add_or_update_crop(crop_data: Dict[str, Any]) -> Dict[str, Any]:
    """신규 작물 등록 또는 기존 작물 정보/단가 수정"""
    name = crop_data.get("name", "").strip()
    if not name:
        raise ValueError("작물명을 입력해주세요.")
    try:
        unit_price = float(crop_data.get("unit_price", 0))
    except (ValueError, TypeError):
        raise ValueError("올바른 단가를 입력해주세요.")

    original_name = crop_data.get("original_name", "").strip()
    category = crop_data.get("category", "일반채소").strip()
    unit = crop_data.get("unit", "g").strip()
    remarks = crop_data.get("remarks", "").strip()

    crops = get_crops_detailed()
    new_crops = []
    found = False

    target_name = original_name if original_name else name
    for c in crops:
        if c["name"] == target_name:
            new_crops.append({
                "name": name,
                "unit_price": unit_price,
                "category": category,
                "unit": unit,
                "remarks": remarks
            })
            found = True
        else:
            new_crops.append(c)

    if not found:
        new_crops.append({
            "name": name,
            "unit_price": unit_price,
            "category": category,
            "unit": unit,
            "remarks": remarks
        })

    save_crops_list(new_crops)
    return {"success": True, "message": f"'{name}' 작물 품목이 성공적으로 저장되었습니다."}

def delete_crop(crop_name: str) -> Dict[str, Any]:
    """작물 품목 삭제"""
    crop_name = crop_name.strip()
    crops = get_crops_detailed()
    new_crops = [c for c in crops if c["name"] != crop_name]
    if len(new_crops) == len(crops):
        raise ValueError("해당 작물을 찾을 수 없습니다.")
    save_crops_list(new_crops)
    return {"success": True, "message": f"'{crop_name}' 작물이 삭제되었습니다."}

def number_to_korean(num: int) -> str:
    """숫자를 한글 금액 표기로 변환 (예: 34500 -> 삼만 사천오백 원정)"""
    if not num or num == 0:
        return "영 원정"
    
    digits = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
    units = ["", "십", "백", "천"]
    big_units = ["", "만", "억", "조"]
    
    is_negative = num < 0
    num = abs(int(num))
    
    result = []
    chunk_idx = 0
    
    while num > 0:
        chunk = num % 10000
        if chunk > 0:
            chunk_str = ""
            for i in range(4):
                digit = (chunk // (10 ** i)) % 10
                if digit > 0:
                    unit = units[i]
                    d_str = digits[digit]
                    chunk_str = d_str + unit + chunk_str
            chunk_str += big_units[chunk_idx]
            result.insert(0, chunk_str)
        num //= 10000
        chunk_idx += 1
        
    res_text = " ".join(result) + " 원정"
    if is_negative:
        res_text = "마이너스 " + res_text
    return res_text

def ensure_excel_exists():
    """엑셀 파일이 없을 경우 기본 시트 구조를 자동으로 생성"""
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws_d = wb.active
        ws_d.title = "납품대장"
        ws_d.append(["연번", "주문일자", "납품처", "품목", "단가", "수량", "금액", "배송일자", "비고", "수령확인", "작물명", "기본단가"])
        default_crops = [("루꼴라", 17.7), ("엽채류", 11.5), ("애플민트", 56.6), ("바질", 0.0), ("딸기", 50.0)]
        for idx, (c, p) in enumerate(default_crops, start=2):
            ws_d.cell(idx, 11, c)
            ws_d.cell(idx, 12, p)
            
        ws_m = wb.create_sheet("협력매장 목록")
        ws_m.append(["매장명", "주소"])
        default_stores = [
            ("재간정", "서울 강북구 4.19로 37"),
            ("온터", "서울 강북구 4.19로 43"),
            ("비르케", "서울 강북구 4.19로 55"),
            ("민지네", "서울 강북구 4.19로 65")
        ]
        for s, a in default_stores:
            ws_m.append([s, a])
        wb.save(EXCEL_PATH)

def get_stores_and_crops() -> Dict[str, Any]:
    ensure_excel_exists()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    
    # 1. Stores
    stores = {}
    if "협력매장 목록" in wb.sheetnames:
        ws_m = wb["협력매장 목록"]
        for r in range(2, ws_m.max_row + 1):
            name = ws_m.cell(r, 1).value
            addr = ws_m.cell(r, 2).value
            if name:
                stores[str(name).strip()] = str(addr).strip() if addr else ""
                
    ws_d = wb["납품대장"]
    for r in range(2, ws_d.max_row + 1):
        name = ws_d.cell(r, 3).value
        if name and str(name).strip() not in stores:
            stores[str(name).strip()] = ""
            
    # Default addresses for well-known stores if missing
    default_addrs = {
        "재간정": "서울 강북구 4.19로 37",
        "온터": "서울 강북구 4.19로 43",
        "비르케": "서울 강북구 4.19로 55",
        "민지네": "서울 강북구 4.19로 65"
    }
    for k, v in default_addrs.items():
        if k in stores and not stores[k]:
            stores[k] = v

    # 2. Crops & default prices
    crops_list = get_crops_detailed()
    crops = {c["name"]: float(c.get("unit_price", 0)) for c in crops_list}

    return {
        "stores": [{"name": k, "address": v} for k, v in stores.items()],
        "crops": crops,
        "crops_detailed": crops_list,
        "supplier": get_supplier_info()
    }

def get_grouped_orders(search_date: Optional[str] = None, search_store: Optional[str] = None, year_month: Optional[str] = None, limit: int = 500) -> List[Dict[str, Any]]:
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["납품대장"]
    
    # Store addresses map
    stores_info = get_stores_and_crops()["stores"]
    store_addr_map = {s["name"]: s["address"] for s in stores_info}

    # Crops default price map for fallback calculation (when Excel has formula or uncomputed values)
    crops_detailed = get_crops_detailed()
    crop_price_map = {c["name"]: float(c.get("unit_price", 0)) for c in crops_detailed}
    groups = {}  # key: (date_str, store_name)

    # 1. 로컬 엑셀 기본 읽기
    raw_rows = []
    for r in range(2, ws.max_row + 1):
        raw_rows.append({
            "row": r,
            "no": ws.cell(r, 1).value,
            "order_date": ws.cell(r, 2).value,
            "store_name": ws.cell(r, 3).value,
            "crop": ws.cell(r, 4).value,
            "unit_price": ws.cell(r, 5).value,
            "qty": ws.cell(r, 6).value,
            "amount": ws.cell(r, 7).value,
            "delivery_date": ws.cell(r, 8).value,
            "remarks": ws.cell(r, 9).value
        })

    # 2. 구글 스프레드시트 실시간 데이터 병합 (웹에서 추가 등록된 최신 건 자동 반영)
    try:
        import google_sheets_manager
        sheet_rows = google_sheets_manager.fetch_ledger_rows_from_sheet()
        if sheet_rows and len(sheet_rows) >= len(raw_rows):
            raw_rows = sheet_rows
    except Exception:
        pass

    for row_obj in raw_rows:
        no = row_obj.get("no")
        order_date_raw = row_obj.get("order_date")
        store_name = row_obj.get("store_name")
        crop = row_obj.get("crop")
        unit_price = row_obj.get("unit_price")
        qty = row_obj.get("qty")
        amount = row_obj.get("amount")
        delivery_date_raw = row_obj.get("delivery_date")
        remarks = row_obj.get("remarks")
        r = row_obj.get("row", 0)
        
        if not order_date_raw and not store_name:
            continue
            
        def normalize_date(val):
            if not val: return ""
            s = str(val).strip().split('(')[0].strip()
            if s.count('/') == 2:
                p = s.split('/')
                if len(p[0]) == 2: p[0] = '20' + p[0]
                return f"{p[0]}-{p[1].zfill(2)}-{p[2].zfill(2)}"
            return s[:10]

        order_date_str = normalize_date(order_date_raw)
        delivery_date_str = normalize_date(delivery_date_raw) or order_date_str
        eff_date_str = delivery_date_str if delivery_date_str else order_date_str
        store_name_str = str(store_name).strip() if store_name else "미지정"
        crop_str = str(crop).strip() if crop else ""
        
        if year_month and not eff_date_str.startswith(year_month) and not order_date_str.startswith(year_month):
            continue
        if search_date and search_date != eff_date_str and search_date != order_date_str:
            continue
        if search_store and search_store not in store_name_str:
            continue

        key = (eff_date_str, store_name_str)
        if key not in groups:
            groups[key] = {
                "order_key": f"{order_date_str}___{store_name_str}",
                "order_date": order_date_str,
                "delivery_date": eff_date_str,
                "store_name": store_name_str,
                "address": store_addr_map.get(store_name_str, ""),
                "items": [],
                "total_qty": 0.0,
                "total_amount": 0.0
            }
            
        qty_val = float(qty) if isinstance(qty, (int, float)) else 0.0
        
        # 단가 계산 (수식이나 None/0일 경우 등록된 작물 단가로 자동 보정)
        if isinstance(unit_price, (int, float)) and float(unit_price) > 0:
            unit_price_val = float(unit_price)
        else:
            unit_price_val = crop_price_map.get(crop_str, 0.0)
            
        # 금액 계산 (수식이나 None/0일 경우 수량 * 단가로 정확히 계산)
        if isinstance(amount, (int, float)) and float(amount) > 0:
            amount_val = float(amount)
        else:
            amount_val = round(qty_val * unit_price_val)
        
        groups[key]["items"].append({
            "row": r,
            "no": no,
            "crop": crop_str,
            "unit_price": unit_price_val,
            "qty": qty_val,
            "amount": amount_val,
            "delivery_date": eff_date_str,
            "remarks": str(remarks) if remarks is not None else ""
        })
        groups[key]["total_qty"] += qty_val
        groups[key]["total_amount"] += amount_val

    # Convert to list and sort by delivery_date descending
    result = list(groups.values())
    result.sort(key=lambda x: (x["delivery_date"] or x["order_date"], x["store_name"]), reverse=True)
    return result[:limit]

def get_receipt_data(order_key: str) -> Optional[Dict[str, Any]]:
    parts = order_key.split("___")
    if len(parts) != 2:
        return None
    order_date_str, store_name_str = parts[0], parts[1]
    
    orders = get_grouped_orders(search_date=order_date_str, search_store=store_name_str, limit=5)
    for ord_data in orders:
        if ord_data["order_date"] == order_date_str and ord_data["store_name"] == store_name_str:
            supplier = get_supplier_info()
            total_amount_int = int(round(ord_data["total_amount"]))
            korean_amount = number_to_korean(total_amount_int)
            return {
                **ord_data,
                "supplier": supplier,
                "korean_amount": korean_amount,
                "formatted_total_amount": f"{total_amount_int:,}",
                "formatted_total_qty": f"{ord_data['total_qty']:,.1f}".rstrip('0').rstrip('.')
            }
    return None

def add_delivery_order(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    주문 저장:
    data = {
        "order_date": "2026-10-09",
        "delivery_date": "2026-10-09",
        "store_name": "까미노",
        "items": [
            {"crop": "루꼴라", "qty": 1000, "unit_price": 17.7, "remarks": ""},
            {"crop": "애플민트", "qty": 300, "unit_price": 56.6, "remarks": ""}
        ]
    }
    """
    order_date = data.get("order_date") or datetime.today().strftime("%Y-%m-%d")
    delivery_date = data.get("delivery_date") or order_date
    store_name = data.get("store_name", "").strip()
    items = data.get("items", [])
    
    if not store_name:
        raise ValueError("매장명을 선택하거나 입력해주세요.")
    if not items:
        raise ValueError("최소 1개 이상의 품목을 입력해주세요.")

    # 1. Backup Excel file
    ensure_excel_exists()
    backup_dir = os.path.join(BASE_DIR, "backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(backup_dir, f"납품대장_{timestamp}.xlsx")
    shutil.copy2(EXCEL_PATH, backup_file)

    # 2. Load workbook
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    ws = wb["납품대장"]

    # 3. Find current maximum serial number (연번) and find append row
    max_no = 0
    append_row = None
    
    for r in range(2, ws.max_row + 1):
        val_no = ws.cell(r, 1).value
        # Check if row is empty
        val_date = ws.cell(r, 2).value
        val_store = ws.cell(r, 3).value
        
        if isinstance(val_no, (int, float)):
            max_no = max(max_no, int(val_no))
        elif isinstance(val_no, str) and val_no.isdigit():
            max_no = max(max_no, int(val_no))
            
        if not val_date and not val_store and append_row is None:
            # First blank row
            append_row = r

    if append_row is None:
        append_row = ws.max_row + 1

    # Date objects
    try:
        dt_order = datetime.strptime(order_date, "%Y-%m-%d")
    except Exception:
        dt_order = order_date

    try:
        dt_delivery = datetime.strptime(delivery_date, "%Y-%m-%d") if delivery_date else dt_order
    except Exception:
        dt_delivery = delivery_date

    # 4. Append items
    for item in items:
        crop = item.get("crop", "").strip()
        qty = float(item.get("qty", 0))
        unit_price = float(item.get("unit_price", 0))
        remarks = item.get("remarks", "")
        amount = round(qty * unit_price)

        max_no += 1
        curr_row = append_row
        append_row += 1

        ws.cell(curr_row, 1, max_no)
        ws.cell(curr_row, 2, dt_order)
        ws.cell(curr_row, 3, store_name)
        ws.cell(curr_row, 4, crop)
        ws.cell(curr_row, 5, unit_price)
        ws.cell(curr_row, 6, qty)
        ws.cell(curr_row, 7, amount)
        ws.cell(curr_row, 8, dt_delivery)
        ws.cell(curr_row, 9, remarks if remarks else None)

    # 5. 구글 스프레드시트 클라우드 실시간 동기화 (설정되어 있을 경우)
    google_synced = False
    try:
        import google_sheets_manager
        if google_sheets_manager.is_google_sheets_configured():
            google_sheets_manager.append_order_to_sheet(data)
            google_synced = True
    except Exception as g_err:
        print("[Google Sheets Sync Warning]:", g_err)

    # 6. 로컬 엑셀 파일 저장 (엑셀 프로그램에서 열려 있을 경우 예외 처리)
    excel_saved = True
    try:
        wb.save(EXCEL_PATH)
    except PermissionError:
        excel_saved = False
        print("[Warning] 납품대장.xlsx 파일이 엑셀에서 열려 있습니다. 백업 폴더에 임시 저장합니다.")
        try:
            wb.save(os.path.join(backup_dir, f"납품대장_미저장_{timestamp}.xlsx"))
        except Exception:
            pass

    order_key = f"{order_date}___{store_name}"
    msg = f"{store_name} ({len(items)}개 품목) 납품 등록 완료!"
    if google_synced:
        msg += " (구글 시트 실시간 연동 완료)"
    if not excel_saved:
        msg += " (※ PC 엑셀 프로그램이 켜져 있어 구글 시트에 우선 기록되었습니다)"

    return {
        "success": True,
        "order_key": order_key,
        "message": msg
    }

def delete_delivery_order(order_key: str) -> Dict[str, Any]:
    """주문 키(날짜___매장명)에 해당하는 납품 내역 행들을 납품대장.xlsx에서 삭제"""
    parts = order_key.split("___")
    if len(parts) != 2:
        raise ValueError("올바르지 않은 주문 키 형식입니다.")
    target_date, target_store = parts[0].strip(), parts[1].strip()

    ensure_excel_exists()
    # 1. 원본 백업
    backup_dir = os.path.join(BASE_DIR, "backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(backup_dir, f"납품대장_삭제전_{timestamp}.xlsx")
    shutil.copy2(EXCEL_PATH, backup_file)

    # 2. 엑셀 로드 및 일치 행 탐색
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    ws = wb["납품대장"]

    matching_rows = []
    for r in range(2, ws.max_row + 1):
        order_date_raw = ws.cell(r, 2).value
        store_name_raw = ws.cell(r, 3).value
        if not order_date_raw and not store_name_raw:
            continue
        order_date_str = str(order_date_raw)[:10] if order_date_raw else ""
        store_name_str = str(store_name_raw).strip() if store_name_raw else ""

        if order_date_str == target_date and store_name_str == target_store:
            matching_rows.append(r)

    if not matching_rows:
        raise ValueError(f"'{target_store}' ({target_date}) 주문 내역을 찾을 수 없습니다.")

    # 3. 역순으로 행 삭제 (인덱스 보존)
    for r in sorted(matching_rows, reverse=True):
        ws.delete_rows(r, 1)

    wb.save(EXCEL_PATH)
    try:
        import threading
        import google_sheets_manager
        if google_sheets_manager.is_google_sheets_configured():
            threading.Thread(target=google_sheets_manager.sync_full_ledger_to_sheet, daemon=True).start()
    except Exception:
        pass

    return {
        "success": True,
        "message": f"'{target_store}' ({target_date}) 납품 내역({len(matching_rows)}개 품목)이 성공적으로 삭제되었습니다."
    }

def update_delivery_order(order_key: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """기존 주문 내역 수정: 기존 행들을 삭제하고 수정된 내용으로 재입력"""
    parts = order_key.split("___")
    if len(parts) != 2:
        raise ValueError("올바르지 않은 주문 키 형식입니다.")
    target_date, target_store = parts[0].strip(), parts[1].strip()

    new_order_date = data.get("order_date") or target_date
    new_delivery_date = data.get("delivery_date") or new_order_date
    new_store_name = data.get("store_name", "").strip() or target_store
    items = data.get("items", [])

    if not items:
        raise ValueError("최소 1개 이상의 품목을 입력해주세요.")

    ensure_excel_exists()
    # 1. 원본 백업
    backup_dir = os.path.join(BASE_DIR, "backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(backup_dir, f"납품대장_수정전_{timestamp}.xlsx")
    shutil.copy2(EXCEL_PATH, backup_file)

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    ws = wb["납품대장"]

    matching_rows = []
    for r in range(2, ws.max_row + 1):
        order_date_raw = ws.cell(r, 2).value
        store_name_raw = ws.cell(r, 3).value
        if not order_date_raw and not store_name_raw:
            continue
        order_date_str = str(order_date_raw)[:10] if order_date_raw else ""
        store_name_str = str(store_name_raw).strip() if store_name_raw else ""

        if order_date_str == target_date and store_name_str == target_store:
            matching_rows.append(r)

    if not matching_rows:
        raise ValueError(f"수정할 기존 주문 내역('{target_store}', {target_date})을 찾을 수 없습니다.")

    insert_at = min(matching_rows)
    # 기존 행 역순 삭제
    for r in sorted(matching_rows, reverse=True):
        ws.delete_rows(r, 1)

    # 날짜 객체
    try:
        dt_order = datetime.strptime(new_order_date, "%Y-%m-%d")
    except Exception:
        dt_order = new_order_date

    try:
        dt_delivery = datetime.strptime(new_delivery_date, "%Y-%m-%d") if new_delivery_date else dt_order
    except Exception:
        dt_delivery = new_delivery_date

    # 새 품목들 해당 위치에 삽입
    for idx, item in enumerate(items):
        curr_row = insert_at + idx
        ws.insert_rows(curr_row, 1)

        crop = item.get("crop", "").strip()
        qty = float(item.get("qty", 0))
        unit_price = float(item.get("unit_price", 0))
        remarks = item.get("remarks", "")
        amount = round(qty * unit_price)

        ws.cell(curr_row, 1, curr_row - 1)
        ws.cell(curr_row, 2, dt_order)
        ws.cell(curr_row, 3, new_store_name)
        ws.cell(curr_row, 4, crop)
        ws.cell(curr_row, 5, unit_price)
        ws.cell(curr_row, 6, qty)
        ws.cell(curr_row, 7, amount)
        ws.cell(curr_row, 8, dt_delivery)
        ws.cell(curr_row, 9, remarks if remarks else None)

    wb.save(EXCEL_PATH)
    try:
        import threading
        import google_sheets_manager
        if google_sheets_manager.is_google_sheets_configured():
            threading.Thread(target=google_sheets_manager.sync_full_ledger_to_sheet, daemon=True).start()
    except Exception:
        pass

    new_order_key = f"{new_order_date}___{new_store_name}"
    return {
        "success": True,
        "order_key": new_order_key,
        "message": f"'{new_store_name}' ({new_order_date}) 납품 내역({len(items)}개 품목)이 성공적으로 수정되었습니다."
    }

def get_all_stores_detailed() -> List[Dict[str, Any]]:
    """협력매장 전체 목록 조회"""
    ensure_excel_exists()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    stores_map = {}

    if "협력매장 목록" in wb.sheetnames:
        ws_m = wb["협력매장 목록"]
        for r in range(2, ws_m.max_row + 1):
            name = ws_m.cell(r, 1).value
            if name:
                s_name = str(name).strip()
                addr = ws_m.cell(r, 2).value or ""
                menu = ws_m.cell(r, 3).value or ""
                price = ws_m.cell(r, 4).value or ""
                stores_map[s_name] = {
                    "name": s_name,
                    "address": str(addr).strip(),
                    "menu": str(menu).strip(),
                    "price": str(price).strip()
                }

    # 납품대장에서도 매장 추출
    ws_d = wb["납품대장"]
    for r in range(2, ws_d.max_row + 1):
        name = ws_d.cell(r, 3).value
        if name:
            s_name = str(name).strip()
            if s_name not in stores_map:
                stores_map[s_name] = {
                    "name": s_name,
                    "address": "",
                    "menu": "",
                    "price": ""
                }

    # 기본 주소 보정
    default_addrs = {
        "재간정": "서울 강북구 4.19로 37",
        "온터": "서울 강북구 4.19로 43",
        "비르케": "서울 강북구 4.19로 55",
        "민지네": "서울 강북구 4.19로 65",
        "419국수김밥": "서울 강북구 4.19로 31",
        "요요네": "서울 강북구 4.19로 16",
        "까미노": "서울 강북구 4.19로 61",
        "미즐 카페엠": "서울 강북구 4.19로 107",
        "봄타코": "서울 강북구 4.19로 60",
        "크을농": "서울 강북구 4.19로 91",
        "남해횟집": "서울 강북구 4.19로 40",
        "이설함흥냉면": "서울 강북구 4.19로 42",
        "화수브루어리": "서울 강북구 인수봉로 238"
    }
    for k, v in default_addrs.items():
        if k in stores_map and not stores_map[k]["address"]:
            stores_map[k]["address"] = v

    return list(stores_map.values())

def add_or_update_store(store_data: Dict[str, Any]) -> Dict[str, Any]:
    """신규 매장 등록 또는 기존 매장 정보 수정"""
    ensure_excel_exists()
    name = store_data.get("name", "").strip()
    address = store_data.get("address", "").strip()
    menu = store_data.get("menu", "").strip()
    price = store_data.get("price", "").strip()

    if not name:
        raise ValueError("매장명을 입력해주세요.")

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    if "협력매장 목록" not in wb.sheetnames:
        ws_m = wb.create_sheet("협력매장 목록")
        ws_m.append(["업소명", "소재지", "대표메뉴", "가격"])
    else:
        ws_m = wb["협력매장 목록"]

    # 기존 매장 검색
    target_row = None
    for r in range(2, ws_m.max_row + 1):
        c_val = ws_m.cell(r, 1).value
        if c_val and str(c_val).strip() == name:
            target_row = r
            break

    if target_row is None:
        target_row = ws_m.max_row + 1

    ws_m.cell(target_row, 1, name)
    ws_m.cell(target_row, 2, address)
    ws_m.cell(target_row, 3, menu)
    ws_m.cell(target_row, 4, price)

    try:
        wb.save(EXCEL_PATH)
    except PermissionError:
        pass

    return {"success": True, "message": f"매장 '{name}' 정보가 저장되었습니다."}

def get_store_settlement(store_name: str, year_month: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None) -> Dict[str, Any]:
    """특정 매장의 월별 또는 기간별 결재(정산) 내역 상세 및 결재서 데이터 생성"""
    ensure_excel_exists()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["납품대장"]

    store_addr = ""
    for s in get_all_stores_detailed():
        if s["name"] == store_name:
            store_addr = s["address"]
            break

    items_list = []
    total_qty = 0.0
    total_amount = 0

    crops_detailed = get_crops_detailed()
    crop_price_map = {c["name"]: float(c.get("unit_price", 0)) for c in crops_detailed}

    for r in range(2, ws.max_row + 1):
        s_val = ws.cell(r, 3).value
        if not s_val or str(s_val).strip() != store_name:
            continue

        raw_date = ws.cell(r, 2).value
        raw_delivery = ws.cell(r, 8).value
        
        if isinstance(raw_delivery, datetime):
            delivery_date_str = raw_delivery.strftime("%Y-%m-%d")
        else:
            delivery_date_str = str(raw_delivery).strip()[:10] if raw_delivery else ""

        if isinstance(raw_date, datetime):
            order_date_str = raw_date.strftime("%Y-%m-%d")
        else:
            order_date_str = str(raw_date).strip()[:10] if raw_date else ""

        # 실질적 기준일자 = 납품일자 우선
        eff_date = delivery_date_str if delivery_date_str else order_date_str

        # 기간별 필터링
        if start_date and eff_date < start_date and order_date_str < start_date:
            continue
        if end_date and eff_date > end_date and order_date_str > end_date:
            continue
        # 월별 필터링 (기간 지정이 없을 경우)
        if not start_date and not end_date and year_month and year_month != "전체":
            if not eff_date.startswith(year_month) and not order_date_str.startswith(year_month):
                continue

        crop = str(ws.cell(r, 4).value or "").strip()
        
        def _sf(v):
            try:
                return float(v)
            except Exception:
                return 0.0

        def _si(v):
            try:
                return int(round(float(v)))
            except Exception:
                return 0

        unit_price = _sf(ws.cell(r, 5).value)
        qty = _sf(ws.cell(r, 6).value)
        amount = _si(ws.cell(r, 7).value)
        
        if unit_price <= 0:
            unit_price = crop_price_map.get(crop, 0.0)
        if amount <= 0:
            amount = round(qty * unit_price)

        delivery_date = str(ws.cell(r, 8).value or "")[:10]
        remarks = str(ws.cell(r, 9).value or "").strip()

        total_qty += qty
        total_amount += amount

        items_list.append({
            "order_date": date_str,
            "crop": crop,
            "unit_price": unit_price,
            "qty": qty,
            "amount": amount,
            "delivery_date": delivery_date or date_str,
            "remarks": remarks
        })

    # 날짜순 정렬
    items_list.sort(key=lambda x: x["order_date"])

    korean_amount = number_to_korean(total_amount)
    supplier = get_supplier_info()

    if start_date and end_date:
        period_label = f"{start_date} ~ {end_date}"
    elif start_date:
        period_label = f"{start_date} 이후"
    elif end_date:
        period_label = f"{end_date} 이전"
    elif year_month and year_month != "전체":
        period_label = f"{year_month.replace('-', '년 ')}월"
    else:
        period_label = "전체 기간"

    return {
        "store_name": store_name,
        "store_address": store_addr,
        "year_month": period_label,
        "period_label": period_label,
        "items": items_list,
        "total_count": len(items_list),
        "total_qty": round(total_qty, 1),
        "formatted_total_qty": f"{total_qty:,.1f}".rstrip('0').rstrip('.'),
        "total_amount": total_amount,
        "formatted_total_amount": f"{total_amount:,}",
        "korean_amount": korean_amount,
        "supplier": supplier
    }

def get_monthly_summary(year_month: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None) -> Dict[str, Any]:
    """월별 또는 기간별 전체 매장 총괄 정산 집계표"""
    ensure_excel_exists()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["납품대장"]

    stores_data = {}  # store_name -> dict

    crops_detailed = get_crops_detailed()
    crop_price_map = {c["name"]: float(c.get("unit_price", 0)) for c in crops_detailed}

    for r in range(2, ws.max_row + 1):
        s_val = ws.cell(r, 3).value
        if not s_val:
            continue
        store_name = str(s_val).strip()

        raw_date = ws.cell(r, 2).value
        raw_delivery = ws.cell(r, 8).value
        
        if isinstance(raw_delivery, datetime):
            delivery_date_str = raw_delivery.strftime("%Y-%m-%d")
        else:
            delivery_date_str = str(raw_delivery).strip()[:10] if raw_delivery else ""

        if isinstance(raw_date, datetime):
            order_date_str = raw_date.strftime("%Y-%m-%d")
        else:
            order_date_str = str(raw_date).strip()[:10] if raw_date else ""

        eff_date = delivery_date_str if delivery_date_str else order_date_str

        # 기간별 필터링
        if start_date and eff_date < start_date and order_date_str < start_date:
            continue
        if end_date and eff_date > end_date and order_date_str > end_date:
            continue
        # 월별 필터링 (기간 지정이 없을 경우)
        if not start_date and not end_date and year_month and year_month != "전체":
            if not eff_date.startswith(year_month) and not order_date_str.startswith(year_month):
                continue

        crop = str(ws.cell(r, 4).value or "").strip()
        try:
            qty = float(ws.cell(r, 6).value or 0)
        except Exception:
            qty = 0.0
        try:
            amount = int(round(float(ws.cell(r, 7).value or 0)))
        except Exception:
            amount = 0

        if amount <= 0:
            u_p = crop_price_map.get(crop, 0.0)
            amount = round(qty * u_p)

        if store_name not in stores_data:
            stores_data[store_name] = {
                "store_name": store_name,
                "order_count": 0,
                "total_qty": 0.0,
                "total_amount": 0,
                "crops_breakdown": {}
            }

        st = stores_data[store_name]
        st["order_count"] += 1
        st["total_qty"] += qty
        st["total_amount"] += amount
        if crop:
            st["crops_breakdown"][crop] = st["crops_breakdown"].get(crop, 0.0) + qty

    # 리스트 변환 및 통계 계산
    result_list = []
    grand_total_qty = 0.0
    grand_total_amount = 0
    grand_total_orders = 0

    for s_name, data in stores_data.items():
        grand_total_qty += data["total_qty"]
        grand_total_amount += data["total_amount"]
        grand_total_orders += data["order_count"]
        crops_text = ", ".join([f"{c} {q:,.0f}g" for c, q in data["crops_breakdown"].items()])
        result_list.append({
            "store_name": s_name,
            "order_count": data["order_count"],
            "total_qty": data["total_qty"],
            "formatted_total_qty": f"{data['total_qty']:,.1f}".rstrip('0').rstrip('.'),
            "total_amount": data["total_amount"],
            "formatted_total_amount": f"{data['total_amount']:,}",
            "crops_summary": crops_text
        })

    # 금액 내림차순 정렬
    result_list.sort(key=lambda x: x["total_amount"], reverse=True)
    supplier = get_supplier_info()

    if start_date and end_date:
        period_label = f"{start_date} ~ {end_date}"
    elif start_date:
        period_label = f"{start_date} 이후"
    elif end_date:
        period_label = f"{end_date} 이전"
    elif year_month and year_month != "전체":
        period_label = f"{year_month.replace('-', '년 ')}월"
    else:
        period_label = "전체 기간"

    return {
        "year_month": period_label,
        "period_label": period_label,
        "stores": result_list,
        "store_count": len(result_list),
        "grand_total_orders": grand_total_orders,
        "grand_total_qty": grand_total_qty,
        "formatted_grand_total_qty": f"{grand_total_qty:,.1f}".rstrip('0').rstrip('.'),
        "grand_total_amount": grand_total_amount,
        "formatted_grand_total_amount": f"{grand_total_amount:,}",
        "korean_amount": number_to_korean(grand_total_amount),
        "supplier": supplier
    }

def save_uploaded_excel(file_content: bytes) -> Dict[str, Any]:
    """업로드된 최신 엑셀 파일 저장 및 교체"""
    # 1. 파일 검증
    import io
    try:
        test_wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
        if "납품대장" not in test_wb.sheetnames:
            raise ValueError("업로드된 엑셀 파일에 '납품대장' 시트가 없습니다.")
    except Exception as e:
        raise ValueError(f"유효한 엑셀 파일이 아닙니다: {e}")

    # 2. 기존 파일 백업
    backup_dir = os.path.join(BASE_DIR, "backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(backup_dir, f"납품대장_업로드이전_{timestamp}.xlsx")
    if os.path.exists(EXCEL_PATH):
        try:
            shutil.copy2(EXCEL_PATH, backup_path)
        except Exception:
            pass

    # 3. 새 파일 저장
    with open(EXCEL_PATH, "wb") as f:
        f.write(file_content)

    return {
        "success": True,
        "message": f"최신 엑셀 파일이 성공적으로 적용되었습니다. (시트: {', '.join(test_wb.sheetnames)})"
    }
