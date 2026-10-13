"""
강북구 스마트팜 로컬푸드 납품관리 시스템 - 통계 및 리포트 엔진
1. 매장별 정산서 (Store Settlement)
2. 월별/기간별 총괄 정산표 (Monthly/Periodic Summary)
3. 작물별 공급 분석표 (Crop Analytics)
4. 매장별 납품 순위표 (Store Rankings)
5. 일자별 출고 대장 (Daily Dispatch Ledger)
"""
from typing import List, Dict, Any, Optional
from datetime import datetime
import database

def get_store_settlement_report(
    store_name: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    year_month: Optional[str] = None
) -> Dict[str, Any]:
    """1. 매장별 납품 결재/정산 내역서 집계"""
    orders = database.get_orders(search_store=store_name, limit=2000)
    
    filtered_items = []
    total_qty = 0.0
    total_amount = 0
    crop_stats = {}

    for ord_obj in orders:
        if ord_obj.get("store_name") != store_name:
            continue

        o_date = ord_obj.get("order_date") or ""
        d_date = ord_obj.get("delivery_date") or o_date
        eff_date = d_date or o_date

        if start_date and eff_date < start_date:
            continue
        if end_date and eff_date > end_date:
            continue
        if not start_date and not end_date and year_month and year_month != "전체":
            if not eff_date.startswith(year_month):
                continue

        for item in ord_obj.get("items", []):
            crop = item.get("crop", "").strip()
            qty = float(item.get("qty", 0))
            unit_price = float(item.get("unit_price", 0))
            amt = int(round(item.get("amount", qty * unit_price)))

            total_qty += qty
            total_amount += amt
            crop_stats[crop] = crop_stats.get(crop, 0.0) + qty

            filtered_items.append({
                "order_date": o_date,
                "delivery_date": d_date,
                "crop": crop,
                "unit_price": unit_price,
                "qty": qty,
                "amount": amt,
                "remarks": item.get("remarks", "")
            })

    filtered_items.sort(key=lambda x: (x.get("delivery_date") or x.get("order_date"), x.get("crop")))

    store_addr = ""
    store = database.get_store(store_name)
    if store:
        store_addr = store.get("address", "")

    supplier = database.get_supplier_info()
    from delivery_manager import number_to_korean
    korean_amount = number_to_korean(total_amount)

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
        "period_label": period_label,
        "items": filtered_items,
        "total_count": len(filtered_items),
        "total_qty": round(total_qty, 1),
        "formatted_total_qty": f"{total_qty:,.1f}".rstrip('0').rstrip('.'),
        "total_amount": total_amount,
        "formatted_total_amount": f"{total_amount:,}",
        "korean_amount": korean_amount,
        "crop_summary": ", ".join([f"{c} {q:,.0f}g" for c, q in crop_stats.items()]),
        "supplier": supplier
    }

def get_monthly_summary_report(
    year_month: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
) -> Dict[str, Any]:
    """2. 월별/기간별 전체 매장 총괄 정산표"""
    orders = database.get_orders(limit=5000)
    stores_data = {}

    for ord_obj in orders:
        s_name = ord_obj.get("store_name", "").strip()
        if not s_name:
            continue

        o_date = ord_obj.get("order_date") or ""
        d_date = ord_obj.get("delivery_date") or o_date
        eff_date = d_date or o_date

        if start_date and eff_date < start_date:
            continue
        if end_date and eff_date > end_date:
            continue
        if not start_date and not end_date and year_month and year_month != "전체":
            if not eff_date.startswith(year_month):
                continue

        if s_name not in stores_data:
            stores_data[s_name] = {
                "store_name": s_name,
                "order_count": 0,
                "item_count": 0,
                "total_qty": 0.0,
                "total_amount": 0,
                "crops_breakdown": {}
            }

        st = stores_data[s_name]
        st["order_count"] += 1

        for item in ord_obj.get("items", []):
            crop = item.get("crop", "").strip()
            qty = float(item.get("qty", 0))
            amt = int(round(item.get("amount", 0)))

            st["item_count"] += 1
            st["total_qty"] += qty
            st["total_amount"] += amt
            if crop:
                st["crops_breakdown"][crop] = st["crops_breakdown"].get(crop, 0.0) + qty

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
            "item_count": data["item_count"],
            "total_qty": data["total_qty"],
            "formatted_total_qty": f"{data['total_qty']:,.1f}".rstrip('0').rstrip('.'),
            "total_amount": data["total_amount"],
            "formatted_total_amount": f"{data['total_amount']:,}",
            "crops_summary": crops_text
        })

    # 매출액 내림차순 정렬
    result_list.sort(key=lambda x: x["total_amount"], reverse=True)
    supplier = database.get_supplier_info()
    from delivery_manager import number_to_korean

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

def get_crop_analytics_report(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    year_month: Optional[str] = None
) -> Dict[str, Any]:
    """3. 작물별 공급 분석표 (판매량, 매출액, 점유율, 공급처수)"""
    orders = database.get_orders(limit=5000)
    crop_stats = {}
    grand_total_qty = 0.0
    grand_total_amount = 0

    for ord_obj in orders:
        s_name = ord_obj.get("store_name", "").strip()
        o_date = ord_obj.get("order_date") or ""
        d_date = ord_obj.get("delivery_date") or o_date
        eff_date = d_date or o_date

        if start_date and eff_date < start_date:
            continue
        if end_date and eff_date > end_date:
            continue
        if not start_date and not end_date and year_month and year_month != "전체":
            if not eff_date.startswith(year_month):
                continue

        for item in ord_obj.get("items", []):
            crop = item.get("crop", "").strip() or "미지정 작물"
            qty = float(item.get("qty", 0))
            amt = int(round(item.get("amount", 0)))

            if crop not in crop_stats:
                crop_stats[crop] = {
                    "crop_name": crop,
                    "total_qty": 0.0,
                    "total_amount": 0,
                    "delivery_count": 0,
                    "stores": set()
                }

            cs = crop_stats[crop]
            cs["total_qty"] += qty
            cs["total_amount"] += amt
            cs["delivery_count"] += 1
            if s_name:
                cs["stores"].add(s_name)

            grand_total_qty += qty
            grand_total_amount += amt

    crops_list = []
    for c_name, data in crop_stats.items():
        qty_share = (data["total_qty"] / grand_total_qty * 100) if grand_total_qty > 0 else 0.0
        amt_share = (data["total_amount"] / grand_total_amount * 100) if grand_total_amount > 0 else 0.0
        avg_price = (data["total_amount"] / data["total_qty"]) if data["total_qty"] > 0 else 0.0

        crops_list.append({
            "crop_name": c_name,
            "total_qty": data["total_qty"],
            "formatted_total_qty": f"{data['total_qty']:,.1f}".rstrip('0').rstrip('.'),
            "total_amount": data["total_amount"],
            "formatted_total_amount": f"{data['total_amount']:,}",
            "delivery_count": data["delivery_count"],
            "store_count": len(data["stores"]),
            "stores_summary": ", ".join(sorted(list(data["stores"]))),
            "qty_share": round(qty_share, 1),
            "amt_share": round(amt_share, 1),
            "avg_unit_price": round(avg_price, 1)
        })

    crops_list.sort(key=lambda x: x["total_amount"], reverse=True)

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
        "period_label": period_label,
        "crops": crops_list,
        "crop_count": len(crops_list),
        "grand_total_qty": grand_total_qty,
        "formatted_grand_total_qty": f"{grand_total_qty:,.1f}".rstrip('0').rstrip('.'),
        "grand_total_amount": grand_total_amount,
        "formatted_grand_total_amount": f"{grand_total_amount:,}"
    }

def get_store_ranking_report(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    year_month: Optional[str] = None,
    sort_by: str = "amount"  # 'amount', 'qty', 'count'
) -> Dict[str, Any]:
    """4. 매장별 납품 실적 순위표"""
    summary = get_monthly_summary_report(year_month=year_month, start_date=start_date, end_date=end_date)
    stores = summary.get("stores", [])

    if sort_by == "qty":
        stores.sort(key=lambda x: x["total_qty"], reverse=True)
    elif sort_by == "count":
        stores.sort(key=lambda x: x["order_count"], reverse=True)
    else:
        stores.sort(key=lambda x: x["total_amount"], reverse=True)

    # 랭킹 부여
    for idx, s in enumerate(stores, start=1):
        s["rank"] = idx

    summary["stores"] = stores
    summary["sort_by"] = sort_by
    return summary

def get_daily_dispatch_report(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
) -> Dict[str, Any]:
    """5. 일자별 출고 대장 (시간순/일자별 묶음 배송 내역)"""
    orders = database.get_orders(limit=5000)
    daily_groups = {}
    grand_total_qty = 0.0
    grand_total_amount = 0

    for ord_obj in orders:
        o_date = ord_obj.get("order_date") or ""
        d_date = ord_obj.get("delivery_date") or o_date
        eff_date = d_date or o_date

        if start_date and eff_date < start_date:
            continue
        if end_date and eff_date > end_date:
            continue

        if eff_date not in daily_groups:
            daily_groups[eff_date] = {
                "delivery_date": eff_date,
                "order_count": 0,
                "total_qty": 0.0,
                "total_amount": 0,
                "orders": []
            }

        grp = daily_groups[eff_date]
        grp["order_count"] += 1
        ord_qty = float(ord_obj.get("total_qty", 0))
        ord_amt = int(round(ord_obj.get("total_amount", 0)))

        grp["total_qty"] += ord_qty
        grp["total_amount"] += ord_amt
        grand_total_qty += ord_qty
        grand_total_amount += ord_amt

        crops_desc = ", ".join([f"{i.get('crop')} ({i.get('qty'):,.0f}g)" for i in ord_obj.get("items", [])])
        grp["orders"].append({
            "order_key": ord_obj.get("order_key"),
            "store_name": ord_obj.get("store_name"),
            "crops_desc": crops_desc,
            "total_qty": ord_qty,
            "total_amount": ord_amt,
            "formatted_amount": f"{ord_amt:,}"
        })

    # 날짜 내림차순 정렬
    sorted_days = sorted(daily_groups.values(), key=lambda x: x["delivery_date"], reverse=True)
    for day in sorted_days:
        day["formatted_total_qty"] = f"{day['total_qty']:,.1f}".rstrip('0').rstrip('.')
        day["formatted_total_amount"] = f"{day['total_amount']:,}"

    period_label = f"{start_date} ~ {end_date}" if start_date and end_date else "전체 기간"

    return {
        "period_label": period_label,
        "days": sorted_days,
        "day_count": len(sorted_days),
        "grand_total_qty": grand_total_qty,
        "formatted_grand_total_qty": f"{grand_total_qty:,.1f}".rstrip('0').rstrip('.'),
        "grand_total_amount": grand_total_amount,
        "formatted_grand_total_amount": f"{grand_total_amount:,}"
    }
