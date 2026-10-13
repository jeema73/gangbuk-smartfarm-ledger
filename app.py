from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os
import uvicorn
import database
import reports
import excel_export
import tunnel_manager
import webbrowser
import threading
import time

app = FastAPI(title="강북구청 스마트팜 납품관리 시스템 (Firebase Firestore)")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

class ItemModel(BaseModel):
    crop: str
    qty: float
    unit_price: float
    remarks: Optional[str] = ""

class OrderCreateModel(BaseModel):
    order_date: str
    delivery_date: Optional[str] = None
    store_name: str
    items: List[ItemModel]

class SupplierModel(BaseModel):
    name: str
    department: str
    biz_no: str
    address: str
    manager: str
    tel: str
    delivery_manager: Optional[str] = ""
    delivery_tel: Optional[str] = ""
    signature_image: Optional[str] = ""

@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

# ========================================================
# 1. 메타 데이터 및 기초정보 (매장, 작물, 공급자)
# ========================================================
@app.get("/api/meta")
async def get_meta():
    try:
        stores = database.get_all_stores()
        crops = database.get_all_crops()
        supplier = database.get_supplier_info()
        return JSONResponse(content={
            "stores": stores,
            "crops": crops,
            "supplier": supplier
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/crops")
async def get_crops():
    try:
        crops = database.get_all_crops()
        return JSONResponse(content=crops)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/crops")
async def save_crop(payload: Dict[str, Any]):
    try:
        res = database.add_or_update_crop(payload)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/crops/{crop_name:path}")
async def delete_crop(crop_name: str):
    from urllib.parse import unquote
    crop_name = unquote(crop_name)
    try:
        res = database.delete_crop(crop_name)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/stores")
async def get_stores():
    try:
        stores = database.get_all_stores()
        return JSONResponse(content=stores)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/stores")
async def save_store(payload: Dict[str, Any]):
    try:
        res = database.add_or_update_store(payload)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/supplier")
async def update_supplier(payload: SupplierModel):
    try:
        database.save_supplier_info(payload.dict())
        return JSONResponse(content={"success": True, "message": "공급자 및 담당자 정보가 저장되었습니다."})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ========================================================
# 2. 주문/납품 CRUD
# ========================================================
@app.get("/api/orders")
async def get_orders(search_date: Optional[str] = None, search_store: Optional[str] = None, year_month: Optional[str] = None, limit: int = 500):
    try:
        orders = database.get_orders(search_date=search_date, search_store=search_store, year_month=year_month, limit=limit)
        return JSONResponse(content=orders)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/orders")
async def create_order(payload: OrderCreateModel):
    try:
        result = database.add_delivery_order(payload.dict())
        return JSONResponse(content=result)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.put("/api/orders/{order_key:path}")
async def update_order(order_key: str, payload: OrderCreateModel):
    from urllib.parse import unquote
    order_key = unquote(order_key)
    try:
        result = database.update_delivery_order(order_key, payload.dict())
        return JSONResponse(content=result)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/orders/{order_key:path}")
async def delete_order(order_key: str):
    from urllib.parse import unquote
    order_key = unquote(order_key)
    try:
        result = database.delete_delivery_order(order_key)
        return JSONResponse(content=result)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/receipt/{order_key:path}")
async def get_receipt(order_key: str):
    from urllib.parse import unquote
    order_key = unquote(order_key)
    try:
        receipt = database.get_receipt_data(order_key)
        return JSONResponse(content=receipt)
    except ValueError as ve:
        raise HTTPException(status_code=404, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ========================================================
# 3. 리포트 엔진 (정산서, 월별총괄, 작물분석, 순위, 일자별출고)
# ========================================================
@app.get("/api/reports/store-settlement")
async def report_store_settlement(store_name: str, start_date: Optional[str] = None, end_date: Optional[str] = None, year_month: Optional[str] = None):
    try:
        res = reports.get_store_settlement_report(store_name=store_name, start_date=start_date, end_date=end_date, year_month=year_month)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/reports/monthly-summary")
@app.get("/api/monthly-summary")
async def report_monthly_summary(year_month: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None):
    try:
        res = reports.get_monthly_summary_report(year_month=year_month, start_date=start_date, end_date=end_date)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/reports/crop-analytics")
async def report_crop_analytics(start_date: Optional[str] = None, end_date: Optional[str] = None, year_month: Optional[str] = None):
    try:
        res = reports.get_crop_analytics_report(start_date=start_date, end_date=end_date, year_month=year_month)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/reports/store-ranking")
async def report_store_ranking(start_date: Optional[str] = None, end_date: Optional[str] = None, year_month: Optional[str] = None, sort_by: str = "amount"):
    try:
        res = reports.get_store_ranking_report(start_date=start_date, end_date=end_date, year_month=year_month, sort_by=sort_by)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/reports/daily-dispatch")
async def report_daily_dispatch(start_date: Optional[str] = None, end_date: Optional[str] = None):
    try:
        res = reports.get_daily_dispatch_report(start_date=start_date, end_date=end_date)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ========================================================
# 4. 엑셀 다운로드 (관공서 납품대장 형식 .xlsx 실시간 생성)
# ========================================================
@app.get("/api/excel/download")
async def download_excel(
    search_date: Optional[str] = None,
    search_store: Optional[str] = None,
    year_month: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
):
    try:
        orders = database.get_orders(search_date=search_date, search_store=search_store, year_month=year_month, limit=10000)
        
        # 추가 기간 필터링이 있는 경우
        if start_date or end_date:
            filtered = []
            for o in orders:
                eff = o.get("delivery_date") or o.get("order_date") or ""
                if start_date and eff < start_date:
                    continue
                if end_date and eff > end_date:
                    continue
                filtered.append(o)
            orders = filtered

        suffix_parts = []
        if search_store: suffix_parts.append(search_store)
        if year_month and year_month != "전체": suffix_parts.append(year_month)
        if start_date and end_date: suffix_parts.append(f"{start_date}_{end_date}")
        elif search_date: suffix_parts.append(search_date)

        suffix = "_".join(suffix_parts)
        excel_bytes = excel_export.export_ledger_to_excel(orders, title_suffix=suffix)

        filename = f"납품대장_{suffix}.xlsx" if suffix else f"납품대장_전체.xlsx"
        from urllib.parse import quote
        encoded_filename = quote(filename)

        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/tunnel")
async def get_tunnel_status():
    return JSONResponse(content=tunnel_manager.get_tunnel_info())

def open_browser():
    try:
        time.sleep(1.2)
        webbrowser.open("http://127.0.0.1:8000")
    except Exception:
        pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    if os.name == 'nt' and os.path.exists("cloudflared.exe"):
        tunnel_manager.start_tunnel_background(port=port)
        threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")
