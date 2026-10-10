from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os
import uvicorn
import delivery_manager
import tunnel_manager
import webbrowser
import threading
import time

app = FastAPI(title="강북구청 스마트팜 납품관리 시스템")

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

@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/api/meta")
async def get_meta():
    try:
        data = delivery_manager.get_stores_and_crops()
        return JSONResponse(content=data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/orders")
async def get_orders(search_date: Optional[str] = None, search_store: Optional[str] = None, year_month: Optional[str] = None, limit: int = 500):
    try:
        orders = delivery_manager.get_grouped_orders(search_date=search_date, search_store=search_store, year_month=year_month, limit=limit)
        return JSONResponse(content=orders)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/receipt/{order_key:path}")
async def get_receipt(order_key: str):
    from urllib.parse import unquote
    order_key = unquote(order_key)
    try:
        receipt = delivery_manager.get_receipt_data(order_key)
        if not receipt:
            raise HTTPException(status_code=404, detail="해당 납품 내역을 찾을 수 없습니다.")
        return JSONResponse(content=receipt)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/orders")
async def create_order(payload: OrderCreateModel):
    try:
        data = payload.dict()
        result = delivery_manager.add_delivery_order(data)
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
        result = delivery_manager.delete_delivery_order(order_key)
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
        data = payload.dict()
        result = delivery_manager.update_delivery_order(order_key, data)
        return JSONResponse(content=result)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/crops")
async def get_crops():
    try:
        crops = delivery_manager.get_crops_detailed()
        return JSONResponse(content=crops)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/crops")
async def save_crop(payload: Dict[str, Any]):
    try:
        res = delivery_manager.add_or_update_crop(payload)
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
        res = delivery_manager.delete_crop(crop_name)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/supplier")
async def update_supplier(payload: SupplierModel):
    try:
        delivery_manager.save_supplier_info(payload.dict())
        return JSONResponse(content={"success": True, "message": "공급자 정보가 저장되었습니다."})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/tunnel")
async def get_tunnel_status():
    return JSONResponse(content=tunnel_manager.get_tunnel_info())

@app.get("/api/google-sheet")
async def get_google_sheet_info():
    import google_sheets_manager
    return JSONResponse(content={
        "configured": google_sheets_manager.is_google_sheets_configured(),
        "config": google_sheets_manager.get_google_sheet_config()
    })

@app.post("/api/google-sheet")
async def update_google_sheet_info(payload: Dict[str, Any]):
    import google_sheets_manager
    google_sheets_manager.save_google_sheet_config(payload)
    return JSONResponse(content={"success": True, "message": "구글 시트 연동 설정이 저장되었습니다."})

@app.post("/api/google-sheet/sync-now")
async def sync_google_sheet_now():
    import google_sheets_manager
    try:
        res = google_sheets_manager.sync_full_ledger_to_sheet()
        if not res.get("success"):
            raise HTTPException(status_code=400, detail=res.get("message"))
        return JSONResponse(content=res)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/stores")
async def get_stores():
    try:
        stores = delivery_manager.get_all_stores_detailed()
        return JSONResponse(content=stores)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/stores")
async def save_store(payload: Dict[str, Any]):
    try:
        res = delivery_manager.add_or_update_store(payload)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/settlement")
async def get_settlement(store_name: str, year_month: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None):
    try:
        res = delivery_manager.get_store_settlement(store_name, year_month=year_month, start_date=start_date, end_date=end_date)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/monthly-summary")
async def get_monthly_summary(year_month: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None):
    try:
        res = delivery_manager.get_monthly_summary(year_month=year_month, start_date=start_date, end_date=end_date)
        return JSONResponse(content=res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        res = delivery_manager.save_uploaded_excel(content)
        return JSONResponse(content=res)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

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
