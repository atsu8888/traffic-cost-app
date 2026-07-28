# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（NAVITIMEフル稼働・最終版）

機能:
  - Gemini(Web検索)による高精度な施設名・住所ジオコーディング
  - 【修正】NAVITIMEの飛行機制限を完全解除し、全国どこでも最速ルートを検索
  - 【修正】満席エラーを回避するため、検索日時を「3週間後」に設定
  - NAVITIMEが万が一エラーになった場合のみAIの推測データをフェールセーフとして使用
  - 現地移動（タクシー vs レンタカー 1日12,000円）を比較し最安パターンを1つだけ出力
  - 試算結果を詳細ルート付きでExcelファイル(.xlsx)として自動出力
"""

import json
import math
import time
import io
import re
import requests
from datetime import datetime, timedelta
import streamlit as st
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 定数・固定値
# ============================================================

STATION_COORDS = {
    "大宮駅": (35.906702, 139.623596),
    "淀屋橋駅": (34.693729, 135.499814),
}

HOTEL_COST_PER_NIGHT_PER_PERSON = 20_000   # 宿泊費（1泊/1名）
RENTAL_CAR_COST_PER_DAY = 12_000           # レンタカー（1日）
TAXI_FARE_PER_KM = 400                     # タクシー単価（円/km）

NAVITIME_URL = "https://navitime-route-totalnavi.p.rapidapi.com/route_transit"
NAVITIME_HOST = "navitime-route-totalnavi.p.rapidapi.com"
GSI_GEOCODE_URL = "https://msearch.gsi.go.jp/address-search/AddressSearch"


# ============================================================
# ユーティリティ関数
# ============================================================

def round_up_1000(amount: float) -> int:
    if amount <= 0: return 0
    return int(math.ceil(amount / 1000.0) * 1000)

def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2)
    return 2 * r * math.asin(math.sqrt(a))

def get_active_gemini_model_name():
    try:
        models = [m.name for m in genai.list_models() if "generateContent" in m.supported_generation_methods]
        for target in ["2.5-flash", "2.0-flash", "flash-latest", "1.5-flash"]:
            for name in models:
                if target in name: return name
        if models: return models[0]
    except Exception:
        pass
    return "models/gemini-2.5-flash"

def guess_airport_from_address(address: str) -> str:
    """住所から離島の空港名を推測するバックアップ処理"""
    if "沖永良部" in address or "知名町" in address or "和泊町" in address: return "沖永良部空港"
    if "奄美" in address or "龍郷町" in address: return "奄美空港"
    if "徳之島" in address or "伊仙町" in address or "天城町" in address: return "徳之島空港"
    if "与論" in address: return "与論空港"
    if "石垣" in address: return "新石垣空港"
    if "宮古" in address: return "宮古空港"
    if "久米島" in address: return "久米島空港"
    if "屋久島" in address: return "屋久島空港"
    if "種子島" in address: return "種子島空港"
    if "喜界" in address: return "喜界空港"
    if "沖縄" in address or "那覇" in address or "宜野湾" in address or "浦添" in address: return "那覇空港"
    return "最寄り空港"


# ============================================================
# Excelファイル出力関数
# ============================================================

def create_excel_report(pattern_data: dict, address: str, headcount: int, work_days: int) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交通費見積"
    ws.views.sheetView[0].showGridLines = True

    table_hdr_fill = PatternFill(start_color="D9531E", end_color="D9531E", fill_type="solid")
    dark_gray_fill = PatternFill(start_color="7F7F7F", end_color="7F7F7F", fill_type="solid")
    light_yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    font_title = Font(name="游ゴシック", size=11, bold=True)
    font_bold = Font(name="游ゴシック", size=9, bold=True)
    font_normal = Font(name="游ゴシック", size=9)
    font_hdr_white = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    ws["A1"] = "設置設定作業"
    ws["A1"].font = font_title
    ws["B4"] = "住所"
    ws["B4"].font = font_title
    ws["B4"].fill = light_yellow_fill
    ws["B4"].alignment = align_center
    ws["C4"] = f"〒{address}"
    ws["C4"].font = font_title
    ws["C4"].fill = light_yellow_fill
    ws["B5"] = "人数"
    ws["B5"].font = font_bold
    ws["B5"].alignment = align_center
    ws["C5"] = f"{headcount} 人"
    ws["C5"].font = font_normal
    ws["C5"].alignment = align_center
    ws["B6"] = "作業"
    ws["B6"].font = font_bold
    ws["B6"].alignment = align_center
    ws["C6"] = f"{work_days} 日"
    ws["C6"].font = font_normal
    ws["C6"].alignment = align_center

    travel_days = 2 if ("前泊" in pattern_data.get("note", "") or "飛行機" in pattern_data.get("name", "")) else 1
    ws["B7"] = "移動"
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center
    ws["C7"] = f"{travel_days} 日"
    ws["C7"].font = font_normal
    ws["C7"].alignment = align_center

    headers = ["チェック", "項目", "経路", "調整費用", "実費用", "日数・泊数", "人数", "合計", "片道移動時間 (h)"]
    for col_num, hdr in enumerate(headers, 1):
        cell = ws.cell(row=10, column=col_num)
        cell.value = hdr
        cell.font = font_hdr_white
        cell.fill = table_hdr_fill
        cell.alignment = align_center
        cell.border = thin_border

    row_idx = 11
    total_sum = 0
    breakdown = pattern_data.get("breakdown", {})
    routes = pattern_data.get("routes", {})

    items_def = [
        {"name": "電車・新幹線（往復）", "key": "電車・新幹線運賃", "route": routes.get("train", "-")},
        {"name": "飛行機（往復）", "key": "航空券費用", "route": routes.get("flight", "-")},
        {"name": "電車・新幹線（往復）", "key": "アクセス電車運賃", "route": routes.get("access", "-")},
        {"name": "レンタカー", "key": "レンタカー費用", "route": routes.get("rental", "-")},
        {"name": "タクシー（往復）", "key": "タクシー運賃", "route": routes.get("taxi", "-")},
        {"name": "宿泊", "key": "宿泊費", "route": "-"},
    ]

    time_hours = round(pattern_data.get("time_min", 0) / 60.0, 1)

    for item in items_def:
        amount = 0
        for k, v in breakdown.items():
            if item["key"] in k:
                amount = v
                break

        is_checked = amount > 0
        chk_mark = "✓" if is_checked else "-"

        ws.cell(row=row_idx, column=1, value=chk_mark).alignment = align_center
        ws.cell(row=row_idx, column=2, value=item["name"]).alignment = align_left
        ws.cell(row=row_idx, column=3, value=item["route"] if is_checked else "-").alignment = align_left
        ws.cell(row=row_idx, column=4, value=amount).number_format = '#,##0'
        ws.cell(row=row_idx, column=5, value=amount).number_format = '#,##0'
        
        days_val = 1
        if "宿泊" in item["name"]:
            days_val = max(work_days, 1) + (1 if "前泊" in pattern_data.get("note", "") else 0)
        elif "レンタカー" in item["name"]:
            days_val = work_days + (1 if "前泊" in pattern_data.get("note", "") else 0)
        elif "タクシー" in item["name"]:
            days_val = work_days + 1

        ws.cell(row=row_idx, column=6, value=days_val if is_checked else 1).alignment = align_center
        ws.cell(row=row_idx, column=7, value=headcount).alignment = align_center
        
        row_total = amount
        ws.cell(row=row_idx, column=8, value=row_total).number_format = '#,##0'
        
        if row_idx == 11 or (row_idx == 12 and is_checked):
            ws.cell(row=row_idx, column=9, value=time_hours).alignment = align_center

        for c in range(1, 10):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = font_normal
            cell.border = thin_border
            if not is_checked and c != 1:
                cell.fill = dark_gray_fill

        total_sum += row_total
        row_idx += 1

    ws.cell(row=row_idx, column=7, value="合計").alignment = align_center
    ws.cell(row=row_idx, column=7).font = font_bold
    total_cell = ws.cell(row=row_idx, column=8, value=pattern_data.get("cost", total_sum))
    total_cell.number_format = '#,##0'
    total_cell.font = font_bold
    total_cell.border = thin_border
    ws.cell(row=row_idx, column=9, value=time_hours).alignment = align_center
    ws.cell(row=row_idx, column=9).font = font_bold
    ws.cell(row=row_idx, column=9).border = thin_border

    # C列（経路）の幅調整
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            ws.column_dimensions[col_letter].width = 35
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による目的地分析 ＋ 国土地理院バックアップ
# ============================================================

def geocode_fallback(address: str):
    """Geminiが座標取得に失敗した際の確実なバックアップ（国土地理院API）"""
    try:
        res = requests.get(GSI_GEOCODE_URL, params={"q": address}, timeout=5)
        if res.status_code == 200 and res.json():
            lon, lat = res.json()[0]["geometry"]["coordinates"]
            return lat, lon
    except:
        pass
    return None, None

def analyze_destination_with_gemini(raw_address: str, api_key: str, origin_name: str) -> dict:
    genai.configure(api_key=api_key.strip())
    origin_city = "大阪" if "淀屋橋" in origin_name else "東京"

    prompt = f"""
出張旅費算出のためGoogle検索で調査してください。

【検索対象】: {raw_address}
【出発拠点】: {origin_city}

以下の形式のJSONテキストのみを回答してください（解説は一切不要）。
{{
  "normalized_address": "対象の正式な住所",
  "dest_lat": 目的地の緯度(数値),
  "dest_lon": 目的地の経度(数値),
  "is_island_or_remote": 対象が海を渡る完全な離島(沖縄、奄美、石垣など)ならtrue。北海道や本州等はfalse,
  "nearest_airport_name": "最寄り空港名(具体名必須。同じ県でも島が違う場合はその島の空港。例:石垣市なら新石垣空港)",
  "airport_lat": 空港の緯度(数値),
  "airport_lon": 空港の経度(数値),
  "flight_fare_estimate": {origin_city}から最寄り空港への大人片道普通運賃概算(数値),
  "flight_time_min": {origin_city}から最寄り空港までの片道総所要時間・分(数値)
}}
"""
    model_name = get_active_gemini_model_name()
    
    fallback_data = {
        "normalized_address": raw_address,
        "is_island_or_remote": "沖縄" in raw_address or "奄美" in raw_address or "沖永良部" in raw_address or "石垣" in raw_address,
        "nearest_airport_name": guess_airport_from_address(raw_address),
        "flight_fare_estimate": 60000,
        "flight_time_min": 180,
        "dest_lat": None, "dest_lon": None,
        "airport_lat": None, "airport_lon": None,
    }

    try:
        model = genai.GenerativeModel(model_name, tools=[{"google_search": {}}])
        res = model.generate_content(prompt)
        text = res.text
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            parsed_data = json.loads(match.group(0))
            fallback_data.update(parsed_data)
    except Exception as e:
        pass

    guessed_airport = guess_airport_from_address(fallback_data["normalized_address"])
    
    if fallback_data["nearest_airport_name"] == "最寄り空港" or not fallback_data["nearest_airport_name"]:
        fallback_data["nearest_airport_name"] = guessed_airport
    elif guessed_airport != "最寄り空港" and guessed_airport != "那覇空港" and fallback_data["nearest_airport_name"] == "那覇空港":
        fallback_data["nearest_airport_name"] = guessed_airport

    if not fallback_data.get("dest_lat") or not fallback_data.get("dest_lon"):
        lat, lon = geocode_fallback(fallback_data["normalized_address"])
        fallback_data["dest_lat"] = lat
        fallback_data["dest_lon"] = lon
        
    if not fallback_data.get("airport_lat"):
        fallback_data["airport_lat"] = fallback_data["dest_lat"] if fallback_data["dest_lat"] else 24.3964
        fallback_data["airport_lon"] = fallback_data["dest_lon"] if fallback_data["dest_lon"] else 124.2450

    return fallback_data


# ============================================================
# STEP 2: NAVITIME API (全国・全交通手段を検索！)
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    """
    飛行機制限を解除し、NAVITIMEの標準アルゴリズムで「最速ルート」を探す。
    直前満席エラーを防ぐため、検索日時を3週間後の平日に設定。
    """
    clean_key = api_key.strip()
    headers = {"X-RapidAPI-Key": clean_key, "X-RapidAPI-Host": NAVITIME_HOST}
    
    # ★ 3週間後の朝9時で検索（満席エラー回避）
    future_day = datetime.now() + timedelta(days=21)
    start_time_iso = future_day.strftime("%Y-%m-%dT09:00:00")
    
    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "start_time": start_time_iso,
        "format": "json"
        # airplaneパラメータを指定しないことで飛行機も新幹線も自由に使わせる
    }

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
        if res.status_code != 200: return None
        
        data = res.json()
        items = data.get("items", [])
        if not items: return None

        # 最速ルートの選定
        fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
        time_min = fastest_item.get("summary", {}).get("move", {}).get("time", 0)
        
        move_info = fastest_item.get("summary", {}).get("move", {})
        total_fare = move_info.get("fare", {}).get("unit_0", 0) if isinstance(move_info.get("fare"), dict) else 0

        walk_time_min = 0
        has_flight = False
        flight_fare = 0
        
        start_station_name = "出発駅"
        end_station_name = "到着駅"
        flight_start_name = None
        flight_end_name = None

        sections = fastest_item.get("sections", [])
        
        # 駅・空港名の抽出ループ
        for i, sec in enumerate(sections):
            m_type = sec.get("move", "")
            sec_type = sec.get("type", "")

            if sec_type == "move" and m_type == "walk":
                walk_time_min += sec.get("time", 0)
                
            if sec_type == "point":
                if start_station_name == "出発駅":
                    start_station_name = sec.get("name", "出発駅")
                end_station_name = sec.get("name", "到着駅")

            if sec_type == "move" and m_type in ["plane", "flight", "air", "airplane", "aeroplane"]:
                has_flight = True
                if "fare" in sec and isinstance(sec["fare"], dict):
                    flight_fare += sec["fare"].get("unit_0", 0)
                    
                # 飛行機の出発・到着空港を取得
                for j in range(i-1, -1, -1):
                    if sections[j].get("type") == "point":
                        flight_start_name = sections[j].get("name", "出発空港")
                        break
                for j in range(i+1, len(sections)):
                    if sections[j].get("type") == "point":
                        flight_end_name = sections[j].get("name", "到着空港")
                        break

        # 運賃の整理
        if has_flight:
            if flight_fare == 0:
                flight_fare = int(total_fare * 0.8)
            access_train_fare = int(total_fare - flight_fare)
        else:
            access_train_fare = 0

        return {
            "has_flight": has_flight,
            "time_min": time_min,
            "total_fare": int(total_fare),
            "flight_fare": int(flight_fare),
            "access_train_fare": int(access_train_fare),
            "walk_time_min": walk_time_min,
            "start_station": start_station_name,
            "end_station": end_station_name,
            "flight_start": flight_start_name,
            "flight_end": flight_end_name
        }
    except Exception as e:
        st.error(f"NAVITIME通信エラー: {e}")
        return None


# ============================================================
# パターン生成と比較ロジック（結果を一つに絞る）
# ============================================================

def build_best_route_patterns(st_name: str, ai_info: dict, navitime_route: dict, headcount: int, work_days: int):
    patterns = []
    
    # 🌟 NAVITIMEが有効なルートを返してくれたかどうか
    if navitime_route:
        has_flight = navitime_route["has_flight"]
        time_min = navitime_route["time_min"]
        
        if has_flight:
            # NAVITIMEの飛行機データを使用
            selected_mode = "flight"
            flight_fare = navitime_route["flight_fare"]
            access_train_fare = navitime_route["access_train_fare"]
            
            origin_airport = navitime_route["flight_start"] or "出発空港"
            airport_name = navitime_route["flight_end"] or "到着空港"
            end_st = navitime_route["end_station"]
            
            flight_routes = {
                "train": "-",
                "flight": f"{origin_airport} ➔ {airport_name}",
                "access": f"{st_name} ➔ {origin_airport}",
                "taxi": f"{airport_name} ↔ 目的地",
                "rental": f"{airport_name} ➔ 目的地周辺"
            }
            route_dict = flight_routes
            base_taxi_km = max(navitime_route["walk_time_min"] * 0.08, 15.0)
            route_title_prefix = f"✈️ 飛行機ルート ({airport_name}利用)"
            display_route_str = f"{st_name} ➔ {origin_airport} ➔ {airport_name} ➔ 目的地"
            is_ai_fare = False
            
        else:
            # NAVITIMEの新幹線データを使用
            selected_mode = "train"
            end_st = navitime_route["end_station"]
            train_routes = {
                "train": f"{st_name} ➔ {end_st}",
                "flight": "-",
                "access": "-",
                "taxi": f"{end_st} ↔ 目的地",
                "rental": f"{end_st} ➔ 目的地周辺"
            }
            route_dict = train_routes
            base_taxi_km = max(navitime_route["walk_time_min"] * 0.08, 1.5)
            route_title_prefix = f"🚄 新幹線/電車ルート ({end_st}着)"
            display_route_str = f"{st_name} ➔ {end_st} ➔ 目的地"
            is_ai_fare = False
            
    else:
        # 🚨 NAVITIMEがどうしてもエラーを返した場合のみ、AIデータで完全にフェールセーフ
        selected_mode = "flight"
        airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
        airport_lat, airport_lon = ai_info.get("airport_lat", 0), ai_info.get("airport_lon", 0)
        dest_lat, dest_lon = ai_info.get("dest_lat", 0), ai_info.get("dest_lon", 0)
        
        airport_to_dest_km = 15.0
        if airport_lat and airport_lon and dest_lat and dest_lon:
            airport_to_dest_km = max(haversine_km(airport_lat, airport_lon, dest_lat, dest_lon), 1.0)
        
        time_min = int(60 + ai_info.get("flight_time_min", 150) + (airport_to_dest_km / 40.0 * 60))
        flight_fare = int(ai_info.get("flight_fare_estimate", 60000) * 1.3) # 1.3倍マージン
        
        origin_airport = "伊丹空港" if "淀屋橋" in st_name else "羽田空港"
        access_train_fare = 1500 if "大宮" in st_name else 500
        
        route_dict = {
            "train": "-",
            "flight": f"{origin_airport} ➔ {airport_name}",
            "access": f"{st_name} ➔ {origin_airport}",
            "taxi": f"{airport_name} ↔ 目的地",
            "rental": f"{airport_name} ➔ 目的地周辺"
        }
        base_taxi_km = airport_to_dest_km
        route_title_prefix = f"✈️ 飛行機ルート ({airport_name}利用)"
        display_route_str = f"{st_name} ➔ {origin_airport} ➔ {airport_name} ➔ 目的地"
        is_ai_fare = True

    travel_hours = time_min / 60.0

    if selected_mode == "flight" or travel_hours >= 4.0:
        nights = work_days + 1
        stay_note = f"移動時間や飛行機利用のため、前泊/後泊想定（{nights}泊）"
    elif travel_hours >= 2.5:
        nights = work_days
        stay_note = f"移動2.5時間以上の為宿泊想定（{nights}泊）"
    else:
        nights = max(work_days - 1, 0)
        stay_note = f"日帰り/標準宿泊（{nights}泊）"

    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)
    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    rental_car_total = round_up_1000(RENTAL_CAR_COST_PER_DAY * rental_days)

    taxi_one_way = base_taxi_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1
    taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

    # タクシーとレンタカーの計算
    b_taxi = {}
    b_rental = {}
    
    if selected_mode == "flight":
        b_taxi["航空券費用(往復・人数分)"] = round_up_1000(flight_fare * 2 * headcount)
        b_taxi["最寄り空港アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
        b_rental["航空券費用(往復・人数分)"] = round_up_1000(flight_fare * 2 * headcount)
        b_rental["最寄り空港アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
    else:
        b_taxi["電車・新幹線運賃(往復・人数分)"] = round_up_1000(navitime_route["total_fare"] * 2 * headcount)
        b_rental["電車・新幹線運賃(往復・人数分)"] = round_up_1000(navitime_route["total_fare"] * 2 * headcount)
        
    b_taxi[f"現地タクシー運賃(往復×{taxi_trips}回分)"] = taxi_total
    b_taxi["宿泊費"] = hotel_cost
    taxi_sum = sum(b_taxi.values())

    b_rental[f"レンタカー費用(12,000円×{rental_days}日)"] = rental_car_total
    b_rental["宿泊費"] = hotel_cost
    rental_sum = sum(b_rental.values())

    # 比較して安い方だけを採用する
    if rental_sum < taxi_sum:
        diff = taxi_sum - rental_sum
        recommend_msg = f"🏆 最安ルート (レンタカー利用・タクシーより {diff:,} 円お得)"
        final_breakdown = b_rental
        final_cost = rental_sum
        final_type = "rental"
        final_name = f"{route_title_prefix} ＋ 現地レンタカー ({rental_days}日間)"
    else:
        recommend_msg = "🏆 最安ルート (タクシー利用推奨)"
        final_breakdown = b_taxi
        final_cost = taxi_sum
        final_type = "taxi"
        final_name = f"{route_title_prefix} ＋ 現地タクシー"

    if is_ai_fare:
        recommend_msg += " / ⚠️ AI相場検索(1.3倍マージン)"

    patterns.append({
        "type": final_type, 
        "name": final_name,
        "time_min": time_min, 
        "cost": final_cost,
        "breakdown": final_breakdown, 
        "note": stay_note,
        "routes": route_dict, 
        "display_route": display_route_str,
        "recommend_reason": recommend_msg,
        "is_recommended": True
    })

    return patterns


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗", layout="wide")
st.title("🚗 交通費・出張見積もりアプリ")

st.sidebar.header("🔑 APIキー設定")
gemini_api_key = st.sidebar.text_input("Gemini API Key", type="password")
navitime_api_key = st.sidebar.text_input("NAVITIME API Key (RapidAPI)", type="password")

if not gemini_api_key or not navitime_api_key:
    st.info("👈 サイドバーから Gemini API Key と NAVITIME API Key を入力してください。")
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    address_input = st.text_input("目的地（住所や施設名）", "沖縄県石垣市美崎町14（石垣市役所）")
with col2:
    station_choice = st.selectbox("出発拠点", ["淀屋橋駅", "大宮駅", "両方比較"], index=0)

col_a, col_b = st.columns(2)
with col_a:
    headcount = st.number_input("作業人数（人）", min_value=1, value=1)
with col_b:
    work_days = st.number_input("現地作業日数（日）", min_value=1, value=2)

st.markdown("---")

if st.button("🚀 最速出張見積もりを計算する", type="primary"):

    stations = STATION_COORDS if station_choice == "両方比較" else {station_choice: STATION_COORDS[station_choice]}

    for st_name, (st_lat, st_lon) in stations.items():
        st.markdown(f"### 🚉 出発地: 【{st_name}】")

        with st.spinner(f"🤖 AIが {address_input} を検索・分析中..."):
            ai_info = analyze_destination_with_gemini(address_input, gemini_api_key, st_name)
            
        if not ai_info.get("dest_lat") or not ai_info.get("dest_lon"):
            st.error("❌ 目的地の座標が取得できませんでした。住所を詳しく入力してください。")
            continue

        st.success(f"**📍 検索地点:** {ai_info.get('normalized_address', address_input)}")
        
        with st.spinner(f"🧭 NAVITIMEで全国対応の最速ルートを検索中..."):
            navitime_route = get_navitime_fastest_route(st_lat, st_lon, ai_info["dest_lat"], ai_info["dest_lon"], navitime_api_key)

        patterns = build_best_route_patterns(st_name, ai_info, navitime_route, headcount, work_days)

        for p in patterns:
            tag_text = f"⭐ {p['recommend_reason']}"
            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円  {tag_text}", expanded=p["is_recommended"]):
                st.success(p['recommend_reason'])

                st.write(f"🗺️ **詳細ルート:** {p['display_route']}")
                st.write(f"⏱ **片道所要時間:** 約 {p['time_min']} 分")
                st.write(f"🏨 **宿泊条件:** {p['note']}")
                st.write("**💰 費目別内訳:**")
                for item, amt in p["breakdown"].items():
                    st.write(f"　・ {item}: **{amt:,}** 円")

                excel_bytes = create_excel_report(p, ai_info.get('normalized_address', address_input), headcount, work_days)
                st.download_button(
                    label=f"📥 Excel見積書をダウンロード",
                    data=excel_bytes,
                    file_name=f"交通費見積書_{st_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{st_name}_{p['type']}"
                )
        st.markdown("---")


