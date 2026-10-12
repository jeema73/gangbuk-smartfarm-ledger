import os
import json
import requests
from datetime import datetime
from typing import List, Dict, Any, Optional

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GOOGLE_KEY_PATH = os.path.join(BASE_DIR, "google_key.json")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def get_google_sheet_config() -> Dict[str, Any]:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("google_sheet", {})
        except Exception:
            pass
    return {}

def save_google_sheet_config(conf: Dict[str, Any]):
    current_config = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                current_config = json.load(f)
        except Exception:
            pass
    current_config["google_sheet"] = conf
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(current_config, f, ensure_ascii=False, indent=2)

def is_google_sheets_configured() -> bool:
    conf = get_google_sheet_config()
    # 1. Webhook 방식 (가장 간편한 Apps Script)
    if conf.get("webhook_url"):
        return True
    # 2. Service Account JSON 방식
    if os.path.exists(GOOGLE_KEY_PATH) and (conf.get("sheet_id") or conf.get("sheet_url")):
        return True
    return False

def append_order_to_sheet(order_data: Dict[str, Any]):
    """구글 시트에 주문 데이터를 전송하여 추가합니다."""
    conf = get_google_sheet_config()

    # 1. Webhook (Google Apps Script) 방식 우선 처리
    webhook_url = conf.get("webhook_url")
    if webhook_url:
        try:
            res = requests.post(webhook_url, json=order_data, timeout=8)
            print("[Google Sheets Webhook] 전송 완료:", res.status_code)
            return
        except Exception as e:
            print("[Google Sheets Webhook Error]:", e)

    # 2. Service Account (gspread) 방식 처리
    if os.path.exists(GOOGLE_KEY_PATH):
        try:
            import gspread
            client = gspread.service_account(filename=GOOGLE_KEY_PATH)
            sheet_url = conf.get("sheet_url")
            sheet_id = conf.get("sheet_id")
            ss = client.open_by_url(sheet_url) if sheet_url else client.open_by_key(sheet_id)
            try:
                ws = ss.worksheet("납품대장")
            except Exception:
                ws = ss.sheet1

            records = ws.get_all_values()
            max_no = 0
            for row in records[1:]:
                if row and row[0].isdigit():
                    max_no = max(max_no, int(row[0]))

            order_date = order_data.get("order_date", "")
            delivery_date = order_data.get("delivery_date") or order_date
            store_name = order_data.get("store_name", "")
            items = order_data.get("items", [])

            rows_to_append = []
            for it in items:
                max_no += 1
                crop = it.get("crop", "")
                qty = it.get("qty", 0)
                unit_price = it.get("unit_price", 0)
                amount = round(qty * unit_price)
                remarks = it.get("remarks", "")

                rows_to_append.append([
                    max_no,
                    order_date,
                    store_name,
                    crop,
                    unit_price,
                    qty,
                    amount,
                    delivery_date,
                    remarks
                ])

            if rows_to_append:
                ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
                print("[Google Sheets gspread] 동기화 성공!")
        except Exception as e:
            print("[Google Sheets gspread Error]:", e)

def sync_full_ledger_to_sheet() -> Dict[str, Any]:
    """로컬 납품대장.xlsx의 전체 데이터(A~I열)와 작물/단가(K~L열)를 구글 스프레드시트에 100% 동기화 전송"""
    conf = get_google_sheet_config()
    webhook_url = conf.get("webhook_url")
    if not webhook_url:
        return {"success": False, "message": "구글 시트 Webhook URL이 설정되지 않았습니다."}

    import openpyxl
    import delivery_manager

    excel_path = delivery_manager.EXCEL_PATH
    if not os.path.exists(excel_path):
        return {"success": False, "message": "납품대장.xlsx 파일이 존재하지 않습니다."}

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "납품대장" not in wb.sheetnames:
        return {"success": False, "message": "납품대장 시트를 찾을 수 없습니다."}

    ws = wb["납품대장"]
    crops_detailed = delivery_manager.get_crops_detailed()
    crop_price_map = {c["name"]: float(c.get("unit_price", 0)) for c in crops_detailed}

    rows_data = []
    for r in range(2, ws.max_row + 1):
        order_date_raw = ws.cell(r, 2).value
        store_name = ws.cell(r, 3).value
        if not order_date_raw and not store_name:
            continue

        no = ws.cell(r, 1).value
        crop = ws.cell(r, 4).value
        unit_price = ws.cell(r, 5).value
        qty = ws.cell(r, 6).value
        amount = ws.cell(r, 7).value
        delivery_date_raw = ws.cell(r, 8).value
        remarks = ws.cell(r, 9).value

        order_date_str = str(order_date_raw)[:10] if order_date_raw else ""
        delivery_date_str = str(delivery_date_raw)[:10] if delivery_date_raw else order_date_str
        store_name_str = str(store_name).strip() if store_name else ""
        crop_str = str(crop).strip() if crop else ""

        qty_val = float(qty) if isinstance(qty, (int, float)) else 0.0
        if isinstance(unit_price, (int, float)) and float(unit_price) > 0:
            unit_price_val = float(unit_price)
        else:
            unit_price_val = crop_price_map.get(crop_str, 0.0)

        if isinstance(amount, (int, float)) and float(amount) > 0:
            amount_val = float(amount)
        else:
            amount_val = round(qty_val * unit_price_val)

        no_val = int(no) if isinstance(no, (int, float)) else (len(rows_data) + 1)

        rows_data.append([
            no_val,
            order_date_str,
            store_name_str,
            crop_str,
            unit_price_val,
            qty_val,
            amount_val,
            delivery_date_str,
            str(remarks) if remarks is not None else ""
        ])

    crops_data = [[c["name"], float(c.get("unit_price", 0.0))] for c in crops_detailed]

    payload = {
        "action": "sync_full",
        "rows": rows_data,
        "crops": crops_data
    }

    try:
        res = requests.post(webhook_url, json=payload, timeout=20)
        if res.status_code == 200:
            return {"success": True, "message": f"구글 스프레드시트에 {len(rows_data)}건의 납품 내역과 작물 단가가 실시간 동기화되었습니다.", "count": len(rows_data)}
        else:
            return {"success": False, "message": f"구글 시트 전송 오류 (HTTP {res.status_code})"}
    except Exception as e:
        return {"success": False, "message": f"구글 시트 통신 오류: {str(e)}"}

def fetch_supplier_from_sheet() -> Optional[Dict[str, str]]:
    """구글 시트 납품대장 N:O 열에서 실시간 공급자 정보를 읽어옵니다."""
    import urllib.parse
    import csv
    import io
    conf = get_google_sheet_config()
    sheet_id = conf.get("sheet_id") or "1sqeqGdTVnGn0Rx5p2kjRmTX843dVTxIrJjdKvyflyBw"
    try:
        sheet_name = urllib.parse.quote("납품대장")
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&range=N1:O10&sheet={sheet_name}"
        req = requests.get(url, timeout=5)
        if req.status_code == 200 and req.text.strip():
            reader = csv.reader(io.StringIO(req.text))
            supp_data = {}
            for row in reader:
                if len(row) >= 2:
                    k = str(row[0]).strip()
                    v = str(row[1]).strip()
                    if k and v:
                        supp_data[k] = v
            key_map = {
                "상호": "name",
                "기관명": "name",
                "부서": "department",
                "사업자번호": "biz_no",
                "등록번호": "biz_no",
                "주소": "address",
                "소재지": "address",
                "담당자": "manager",
                "성함": "manager",
                "연락처": "tel",
                "전화번호": "tel"
            }
            parsed = {}
            for raw_k, raw_v in supp_data.items():
                for mk, pk in key_map.items():
                    if mk in raw_k:
                        parsed[pk] = raw_v
                        break
            if len(parsed) >= 2:
                return parsed
    except Exception as e:
        print("[Google Sheets Fetch Supplier Warning]:", e)
    return None

