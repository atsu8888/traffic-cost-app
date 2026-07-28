# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（最終統合版）

機能:
  - Gemini(Web検索)による高精度な施設名・住所ジオコーディング
  - 離島/遠方の自動判定（NAVITIMEエラー回避）
  - 飛行機を利用する場合はAI運賃に「1.3倍の安全マージン」を掛けて出力
  - UIおよびExcelに「出発駅 ➔ 出発空港 ➔ 到着空港 ➔ 目的地」の詳細ルートを記載
  - 現地移動（タクシー vs レンタカー 1日12,000円）を比較し最安パターンを出力
  - 試算結果を社内フォーマットのExcelファイル(.xlsx)として自動出力
"""

import json
import math
import time
import io
import requests
from datetime import datetime, timedelta
import streamlit as st
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 定数・固定値（社内ルール・設定）
# ============================================================

STATION_COORDS = {
    "大宮駅": (35.906702, 139.623596),
    "淀屋橋駅": (34.693729, 135.499814),
}

HOTEL_COST_PER_NIGHT_PER_PERSON = 20_000   # 宿泊費（1泊/1名）
RENTAL_CAR_COST_PER_DAY = 12_000           # レンタカー（1日）
TAXI_FARE_PER_KM = 400                    # タクシー単価（円/km）

# RapidAPI NAVITIME Totalnavi API
NAVITIME_URL = "https://navitime-route-totalnavi.p.rapidapi.com/route_transit"
NAVITIME_HOST = "navitime-route-totalnavi.p.rapidapi.com"


# ============================================================
# ユーティリティ関数
# ============================================================

def round_up_1000(amount: float) -> int:
    """各費目を1,000円単位で切り上げ"""
    if amount <= 0:
        return 0
    return int(math.ceil(amount / 1000.0) * 1000)

def haversine_km(lat1, lon1, lat2, lon2):
    """2点間の直線距離(km)を算出"""
    r = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    )
    return 2 * r * math.asin(math.sqrt(a))

def get_active_gemini_model_name():
    """現在呼び出し可能なGeminiモデル名を自動取得"""
    try:
        models = [m.name for m in genai.list_models() if "generateContent" in m.supported_generation_methods]
        for target in ["2.5-flash", "2.0-flash", "flash-latest", "1.5-flash"]:
            for name in models:
                if target in name:
                    return name
        if models:
            return models[0]
    except Exception:
        pass
    return "models/gemini-2.5-flash"


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

    # C列（経路）の幅を広めに調整
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            ws.column_dimensions[col_letter].width = 30
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による目的地分析 (高精度ジオコーディング + 飛行機判定)
# ============================================================

def analyze_destination_with_gemini(raw_address: str, api_key: str, origin_name: str) -> dict:
    genai.configure(api_key=api_key.strip())
    origin_city = "大阪（伊丹/関空）" if "淀屋橋" in origin_name else "東京（羽田）"

    prompt = f"""
あなたは出張旅費算出アシスタントです。Google検索を利用して正確な情報を調査してください。

【検索対象】: {raw_address}
【出発拠点】: {origin_city}

【調査指示】
1. 検索対象の正式住所と「緯度(dest_lat)・経度(dest_lon)」を特定してください。
2. 出発拠点から対象場所への移動において、海を渡る必要がある離島（沖縄、奄美、沖永良部など）か、または北海道等で「陸路(新幹線)での到達が非現実的か」を判定(is_island_or_remote)してください。
3. その目的地へ行くための「最寄り空港名」と、その空港の「緯度(airport_lat)・経度(airport_lon)」を特定してください。
4. 出発拠点から最寄り空港までの「大人片道普通運賃（ANA/JAL）」の概算額を検索してください。
5. 出発空港から到着空港までの「片道の飛行時間（分）」を特定してください。

以下のJSON形式のみで出力してください。
{{
  "normalized_address": "鹿児島県大島郡知名町瀬利覚2208",
  "dest_lat": 27.3821,
  "dest_lon": 128.6015,
  "is_island_or_remote": true,
  "nearest_airport_name": "沖永良部空港",
  "airport_lat": 27.4258,
  "airport_lon": 128.6586,
  "flight_fare_estimate": 65000,
  "flight_time_min": 150
}}
"""
    model_name = get_active_gemini_model_name()
    try:
        model = genai.GenerativeModel(model_name, tools=[{"google_search": {}}])
        res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
        
        # SyntaxError対策：安全に余分なマークダウン記号を除去
        text = res.text.strip()
        text = text.replace('```json', '')
        text = text.replace('```', '')
        text = text.strip()
            
        return json.loads(text)
    except Exception as e:
        return None


# ============================================================
# STEP 2: NAVITIME API (純粋な「陸路」専用検索)
# ============================================================

def get_navitime_train_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    clean_key = api_key.strip()
    headers = {
        "X-RapidAPI-Key": clean_key,
        "X-RapidAPI-Host": NAVITIME_HOST,
    }
    next_day = datetime.now() + timedelta(days=1)
    start_time_iso = next_day.strftime("%Y-%m-%dT09:00:00")

    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "start_time": start_time_iso,
        "format": "json",
        "airplane": 0  # 飛行機ルートを排除（陸路の最速を探すため）
    }

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
        if res.status_code != 200:
            return None
    except:
        return None

    data = res.json()
    items = data.get("items", [])
    if not items:
        return None

    fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
    summary = fastest_item.get("summary", {})
    move_info = summary.get("move", {})
    time_min = move_info.get("time", 0)

    total_fare = 0
    if "fare" in move_info and isinstance(move_info["fare"], dict):
        total_fare = move_info["fare"].get("unit_0", 0) or move_info["fare"].get("total", 0)

    walk_time_min = 0
    end_station_name = "最寄り駅"

    sections = fastest_item.get("sections", [])
    for i in reversed(range(len(sections))):
        sec = sections[i]
        if sec.get("type") == "move" and sec.get("move") not in ["walk", "car", "taxi"]:
            if i + 1 < len(sections):
                next_sec = sections[i + 1]
                if next_sec.get("type") == "point":
                    name = next_sec.get("name", "")
                    if name:
                        end_station_name = name
            break

    for sec in sections:
        if sec.get("type") == "move" and sec.get("move") == "walk":
            walk_time_min += sec.get("time", 0)

    return {
        "time_min": time_min,
        "fare": int(total_fare),
        "walk_time_min": walk_time_min,
        "end_station": end_station_name
    }


# ============================================================
# パターン生成と比較ロジック
# ============================================================

def build_best_route_patterns(
    st_name: str,
    ai_info: dict,
    train_route: dict,
    headcount: int,
    work_days: int
):
    patterns = []
    is_island = ai_info.get("is_island_or_remote", False)
    
    # --- ✈️ 飛行機ルートのパラメータ構築 ---
    airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
    airport_lat = ai_info.get("airport_lat", 0)
    airport_lon = ai_info.get("airport_lon", 0)
    dest_lat = ai_info.get("dest_lat", 0)
    dest_lon = ai_info.get("dest_lon", 0)
    
    airport_to_dest_km = haversine_km(airport_lat, airport_lon, dest_lat, dest_lon)
    if airport_to_dest_km < 1.0: airport_to_dest_km = 15.0
    
    flight_time_total = int(60 + ai_info.get("flight_time_min", 150) + (airport_to_dest_km / 40.0 * 60))
    
    # 💡 AI(Web検索)運賃は安全マージン1.3倍を掛ける
    raw_flight_fare = ai_info.get("flight_fare_estimate", 60000)
    flight_fare = int(raw_flight_fare * 1.3)
    
    origin_airport = "伊丹空港" if "淀屋橋" in st_name else "羽田空港"
    access_train_fare = 1500 if "大宮" in st_name else 500
    
    # 飛行機用の詳細ルート設定
    flight_routes = {
        "train": "-",
        "flight": f"{origin_airport} ➔ {airport_name}",
        "access": f"{st_name} ➔ {origin_airport}",
        "taxi": f"{airport_name} ↔ 目的地",
        "rental": f"{airport_name} ➔ 目的地周辺"
    }
    
    # --- 🚄 新幹線ルートのパラメータ構築 ---
    train_time_total = train_route["time_min"] if train_route else 99999
    
    # 新幹線用の詳細ルート設定
    if train_route:
        end_st = train_route["end_station"]
        train_routes = {
            "train": f"{st_name} ➔ {end_st}",
            "flight": "-",
            "access": "-",
            "taxi": f"{end_st} ↔ 目的地",
            "rental": f"{end_st} ➔ 目的地周辺"
        }
    else:
        end_st = "最寄り駅"
        train_routes = {}
    
    # --- 最速判定 ---
    if is_island or (flight_time_total < train_time_total):
        selected_mode = "flight"
        base_time = flight_time_total
        route_dict = flight_routes
        base_taxi_km = airport_to_dest_km
        route_title_prefix = f"✈️ 飛行機最速ルート ({airport_name}利用)"
        display_route_str = f"{st_name} ➔ {origin_airport} ➔ {airport_name} ➔ 目的地"
    else:
        selected_mode = "train"
        base_time = train_time_total
        route_dict = train_routes
        base_taxi_km = max(train_route["walk_time_min"] * 0.08, 1.5)
        route_title_prefix = f"🚄 新幹線/電車最速ルート ({end_st}着)"
        display_route_str = f"{st_name} ➔ {end_st} ➔ 目的地"
        
    travel_hours = base_time / 60.0

    # 宿泊判定
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

    # パターン1: タクシー利用
    b_taxi = {}
    if selected_mode == "flight":
        b_taxi["航空券費用(往復・人数分)"] = round_up_1000(flight_fare * 2 * headcount)
        b_taxi["最寄り空港アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
    else:
        b_taxi["電車・新幹線運賃(往復・人数分)"] = round_up_1000(train_route["fare"] * 2 * headcount)
    b_taxi[f"現地タクシー運賃(往復×{taxi_trips}回分)"] = taxi_total
    b_taxi["宿泊費"] = hotel_cost

    patterns.append({
        "type": "taxi", "name": f"{route_title_prefix} ＋ 現地タクシー",
        "time_min": base_time, "cost": sum(b_taxi.values()),
        "breakdown": b_taxi, "note": stay_note,
        "taxi_cost": taxi_total, "rental_cost": rental_car_total, 
        "routes": route_dict, "display_route": display_route_str
    })

    # パターン2: レンタカー利用
    b_rental = {}
    if selected_mode == "flight":
        b_rental["航空券費用(往復・人数分)"] = round_up_1000(flight_fare * 2 * headcount)
        b_rental["最寄り空港アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
    else:
        b_rental["電車・新幹線運賃(往復・人数分)"] = round_up_1000(train_route["fare"] * 2 * headcount)
    b_rental[f"レンタカー費用(12,000円×{rental_days}日)"] = rental_car_total
    b_rental["宿泊費"] = hotel_cost

    patterns.append({
        "type": "rental", "name": f"{route_title_prefix} ＋ 現地レンタカー ({rental_days}日間)",
        "time_min": base_time, "cost": sum(b_rental.values()),
        "breakdown": b_rental, "note": stay_note,
        "taxi_cost": taxi_total, "rental_cost": rental_car_total, 
        "routes": route_dict, "display_route": display_route_str
    })

    min_p = min(patterns, key=lambda x: x["cost"])
    for p in patterns:
        reasons = []
        p["is_recommended"] = (p == min_p)

        if p["type"] == "rental" and p["rental_cost"] < p["taxi_cost"]:
            diff = p["taxi_cost"] - p["rental_cost"]
            reasons.append(f"🚗 レンタカー推奨（タクシーより {diff:,} 円お得）")
        elif p["type"] == "taxi" and p["taxi_cost"] <= p["rental_cost"]:
            reasons.append("🚕 タクシー利用推奨")

        if selected_mode == "flight":
            reasons.append("⚠️ AI検索概算適用 (運賃1.3倍安全マージン)")

        if p == min_p:
            reasons.insert(0, "🏆 最安")

        p["recommend_reason"] = " / ".join(reasons)

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
    address_input = st.text_input("目的地（住所や施設名）", "鹿児島県大島郡知名町瀬利覚2208（沖永良部徳洲会病院）")
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
            
        if not ai_info:
            st.error("目的地情報の取得に失敗しました。")
            continue

        normalized_address = ai_info.get("normalized_address", address_input)
        dest_lat = ai_info.get("dest_lat")
        dest_lon = ai_info.get("dest_lon")
        is_island = ai_info.get("is_island_or_remote", False)

        st.success(f"**📍 検索地点:** {normalized_address}")
        
        train_route = None
        if is_island:
            st.info("💡 **海を渡る離島・遠隔地と判定されました。NAVITIME検索をスキップし、飛行機ルートを適用します。**")
        else:
            with st.spinner(f"🧭 陸路の最速ルートをNAVITIMEで検索中..."):
                train_route = get_navitime_train_route(st_lat, st_lon, dest_lat, dest_lon, navitime_api_key)

        patterns = build_best_route_patterns(st_name, ai_info, train_route, headcount, work_days)

        for p in patterns:
            tag_text = f"⭐ {p['recommend_reason']}"
            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円  {tag_text}", expanded=p["is_recommended"]):
                if p["is_recommended"]:
                    st.success(p['recommend_reason'])

                st.write(f"🗺️ **詳細ルート:** {p['display_route']}")
                st.write(f"⏱ **片道所要時間:** 約 {p['time_min']} 分")
                st.write(f"🏨 **宿泊条件:** {p['note']}")
                st.write("**💰 費目別内訳:**")
                for item, amt in p["breakdown"].items():
                    st.write(f"　・ {item}: **{amt:,}** 円")

                excel_bytes = create_excel_report(p, normalized_address, headcount, work_days)
                st.download_button(
                    label=f"📥 Excel見積書をダウンロード",
                    data=excel_bytes,
                    file_name=f"交通費見積書_{st_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{st_name}_{p['type']}"
                )
        st.markdown("---")


