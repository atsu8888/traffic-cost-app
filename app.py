# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ

機能:
  - NAVITIME APIの検索結果から「最も所要時間が短い（最速）ルート」を自動選定
  - NAVITIMEのレスポンスから「到着駅/空港名」を抽出して画面・Excelに反映
  - 飛行機ルート判定ロジック（NAVITIMEの `plane`, `flight`, `air`, `airplane` に対応）
  - 飛行機ルートの場合もNAVITIME APIの運賃（航空券代＋アクセス電車代）でダイレクト計算
  - 新幹線 vs 飛行機の2パターン並びを廃止し、最速手段のみを1つ選択表示
  - 現地移動（タクシー vs レンタカー 1日12,000円）を比較し最安パターンを出力
  - 費目は社内ルールに従い1,000円単位で切り上げ
  - 試算結果を社内フォーマットのExcelファイル(.xlsx)として自動出力・ダウンロード可能

必要ライブラリ:
  pip install streamlit google-generativeai requests openpyxl
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
TAXI_WALK_THRESHOLD_MIN = 15               # 最寄駅から徒歩15分以上でタクシー検討

# タクシー想定単価（初乗り＋距離算定用目安）
TAXI_FARE_PER_KM = 400                    # 円/km

# RapidAPI NAVITIME Totalnavi API
NAVITIME_URL = "https://navitime-route-totalnavi.p.rapidapi.com/route_transit"
NAVITIME_HOST = "navitime-route-totalnavi.p.rapidapi.com"
GSI_GEOCODE_URL = "https://msearch.gsi.go.jp/address-search/AddressSearch"


# ============================================================
# ユーティリティ関数
# ============================================================

def round_up_1000(amount: float) -> int:
    """各費目を1,000円単位で切り上げ（社内ルール）"""
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
        models = [
            m.name for m in genai.list_models()
            if "generateContent" in m.supported_generation_methods
        ]
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
# Excelファイル出力関数（社内フォーマット再現）
# ============================================================

def create_excel_report(pattern_data: dict, address: str, headcount: int, work_days: int, start_st: str, end_st: str) -> bytes:
    """
    試算結果から社内規定レイアウトに基づくExcelファイルを生成してバイナリで返す
    """
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

    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # タイトル
    ws["A1"] = "設置設定作業"
    ws["A1"].font = font_title

    # 基本情報
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

    items_def = [
        {"name": "電車・新幹線（往復）", "key": "電車・新幹線運賃", "route": f"{start_st} ➔ {end_st}"},
        {"name": "飛行機（往復）", "key": "航空券費用", "route": f"{start_st} ➔ {end_st}"},
        {"name": "電車・新幹線（往復）", "key": "アクセス電車運賃", "route": "駅 ➔ 空港"},
        {"name": "レンタカー", "key": "レンタカー費用", "route": f"{end_st} ➔ 目的地周遊"},
        {"name": "タクシー（往復）", "key": "タクシー運賃", "route": f"{end_st} ↔ 目的地"},
        {"name": "宿泊", "key": "宿泊費", "route": ""},
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

    # 合計行
    ws.cell(row=row_idx, column=7, value="合計").alignment = align_center
    ws.cell(row=row_idx, column=7).font = font_bold
    total_cell = ws.cell(row=row_idx, column=8, value=pattern_data.get("cost", total_sum))
    total_cell.number_format = '#,##0'
    total_cell.font = font_bold
    total_cell.border = thin_border

    ws.cell(row=row_idx, column=9, value=time_hours).alignment = align_center
    ws.cell(row=row_idx, column=9).font = font_bold
    ws.cell(row=row_idx, column=9).border = thin_border

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による住所正規化
# ============================================================

def analyze_and_normalize_with_gemini(raw_address: str, api_key: str) -> dict:
    genai.configure(api_key=api_key.strip())

    prompt = f"""
以下の目的地の住所情報を正規化してください。

住所入力: {raw_address}

以下のJSON形式のみで返答してください。前置きや説明文は不要です。
{{
  "normalized_address": "鹿児島県大島郡知名町瀬利覚2208"
}}
"""

    selected_model_name = get_active_gemini_model_name()

    try:
        model = genai.GenerativeModel(selected_model_name)
        res = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"},
        )
        return json.loads(res.text)
    except Exception:
        return {"normalized_address": raw_address}


# ============================================================
# STEP 2: 国土地理院APIでジオコーディング（無料）
# ============================================================

def geocode_address(address: str):
    try:
        res = requests.get(GSI_GEOCODE_URL, params={"q": address}, timeout=10)
        res.raise_for_status()
        data = res.json()
        if data:
            lon, lat = data[0]["geometry"]["coordinates"]
            return lat, lon
    except Exception:
        pass
    return None


# ============================================================
# STEP 3: NAVITIME API から最速ルートを判定・完全計算
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    """
    NAVITIME Totalnavi APIから全ルート候補を取得し、
    最も所要時間が短い（最速の）ルートを自動選定して、その詳細運賃・駅名を抽出して返す。
    """
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
    }

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
    except requests.exceptions.RequestException as e:
        st.error(f"❌ NAVITIME 通信エラー: {e}")
        return None

    if res.status_code != 200:
        st.error(f"❌ NAVITIME API レスポンスエラー (HTTP {res.status_code})")
        with st.expander("エラー詳細ログを表示"):
            st.code(res.text)
        return None

    data = res.json()
    items = data.get("items", [])

    if not items:
        return None

    # 最も所要時間（time）が短いアイテムを選定（最速ルート）
    fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))

    summary = fastest_item.get("summary", {})
    move_info = summary.get("move", {})
    time_min = move_info.get("time", 0)

    # 全体合計運賃（NAVITIME算出値）
    total_fare = 0
    if "fare" in move_info and isinstance(move_info["fare"], dict):
        total_fare = move_info["fare"].get("unit_0", 0) or move_info["fare"].get("total", 0)

    # ルート内で飛行機を利用しているか・徒歩時間・運賃・駅名の抽出
    has_flight = False
    flight_fare = 0
    walk_time_min = 0
    
    end_station_name = "目的地周辺"

    sections = fastest_item.get("sections", [])
    
    # 到着駅/空港の特定 (最後の移動手段の到着地点)
    for i in reversed(range(len(sections))):
        sec = sections[i]
        if sec.get("type") == "move" and sec.get("move") != "walk":
            if "node" in sec and isinstance(sec["node"], list) and len(sec["node"]) > 1:
                end_node = sec["node"][-1]
                end_station_name = end_node.get("name", "目的地周辺")
            break

    for sec in sections:
        m_type = sec.get("move", "")
        sec_type = sec.get("type", "")

        if sec_type == "move" and m_type == "walk":
            walk_time_min += sec.get("time", 0)

        # 飛行機（plane / flight / air / airplane）が含まれるか判定
        if sec_type == "move" and m_type in ["plane", "flight", "air", "airplane", "aeroplane"]:
            has_flight = True
            if "fare" in sec and isinstance(sec["fare"], dict):
                flight_fare += sec["fare"].get("unit_0", 0)

    # 運賃内訳の整備
    if has_flight:
        if flight_fare == 0:
            flight_fare = int(total_fare * 0.8)
        access_train_fare = int(total_fare - flight_fare)
    else:
        flight_fare = 0
        access_train_fare = 0

    return {
        "has_flight": has_flight,
        "time_min": time_min,
        "total_fare": int(total_fare),
        "flight_fare": int(flight_fare),
        "access_train_fare": int(access_train_fare),
        "walk_time_min": walk_time_min,
        "end_station": end_station_name
    }


# ============================================================
# 最速ルートに基づく現地交通手段比較 & 見積もり作成
# ============================================================

def build_fastest_route_patterns(
    station_name: str,
    station_lat: float,
    station_lon: float,
    dest_lat: float,
    dest_lon: float,
    fastest_route: dict,
    headcount: int,
    work_days: int,
):
    """
    NAVITIME選定の最速ルートをベースにし、
    現地「タクシー」と「レンタカー」の2パターンのみを作成・比較
    """
    patterns = []
    has_flight = fastest_route["has_flight"]
    time_min = fastest_route["time_min"]
    travel_hours = time_min / 60.0
    end_st = fastest_route.get("end_station", "目的地最寄り")

    # 前泊・宿泊日数判定
    if has_flight or travel_hours >= 4.0:
        nights = work_days + 1
        stay_note = f"移動時間（約{travel_hours:.1f}h）または飛行機利用のため、前泊/後泊想定（{nights}泊）"
    elif travel_hours >= 2.5:
        nights = work_days
        stay_note = f"移動2.5時間以上の為宿泊想定（{nights}泊）"
    else:
        nights = max(work_days - 1, 0)
        stay_note = f"日帰り/標準宿泊（{nights}泊）"

    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)

    # 交通費の計算（往復・人数分）
    if has_flight:
        flight_rt_total = round_up_1000(fastest_route["flight_fare"] * 2 * headcount)
        access_train_rt = round_up_1000(fastest_route["access_train_fare"] * 2 * headcount)
        train_rt = 0
        route_title_prefix = f"✈️ {end_st}着 飛行機ルート"
    else:
        flight_rt_total = 0
        access_train_rt = 0
        train_rt = round_up_1000(fastest_route["total_fare"] * 2 * headcount)
        route_title_prefix = f"🚄 {end_st}着 新幹線/電車ルート"

    # 現地移動費用（タクシー vs レンタカー 1日12,000円）
    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    rental_car_total = round_up_1000(RENTAL_CAR_COST_PER_DAY * rental_days)

    est_taxi_km = max(fastest_route["walk_time_min"] * 0.08, 15.0 if has_flight else 2.0)
    taxi_one_way = est_taxi_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1
    taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

    # 1. 最速ルート ＋ タクシー利用
    breakdown_taxi = {}
    if has_flight:
        breakdown_taxi["航空券費用(往復・人数分)"] = flight_rt_total
        breakdown_taxi["最寄り空港アクセス電車運賃(往復・人数分)"] = access_train_rt
    else:
        breakdown_taxi["電車・新幹線運賃(往復・人数分)"] = train_rt
    breakdown_taxi[f"現地タクシー運賃(往復×{taxi_trips}回分)"] = taxi_total
    breakdown_taxi["宿泊費"] = hotel_cost

    patterns.append({
        "type": "taxi",
        "name": f"{route_title_prefix} ＋ 現地タクシー",
        "time_min": time_min,
        "cost": sum(breakdown_taxi.values()),
        "breakdown": breakdown_taxi,
        "note": stay_note,
        "taxi_cost": taxi_total,
        "rental_cost": rental_car_total,
    })

    # 2. 最速ルート ＋ レンタカー利用
    breakdown_rental = {}
    if has_flight:
        breakdown_rental["航空券費用(往復・人数分)"] = flight_rt_total
        breakdown_rental["最寄り空港アクセス電車運賃(往復・人数分)"] = access_train_rt
    else:
        breakdown_rental["電車・新幹線運賃(往復・人数分)"] = train_rt
    breakdown_rental[f"レンタカー費用(12,000円×{rental_days}日)"] = rental_car_total
    breakdown_rental["宿泊費"] = hotel_cost

    patterns.append({
        "type": "rental",
        "name": f"{route_title_prefix} ＋ 現地レンタカー ({rental_days}日間)",
        "time_min": time_min,
        "cost": sum(breakdown_rental.values()),
        "breakdown": breakdown_rental,
        "note": stay_note,
        "taxi_cost": taxi_total,
        "rental_cost": rental_car_total,
    })

    # 推奨理由の自動作成
    min_p = min(patterns, key=lambda x: x["cost"])
    for p in patterns:
        reasons = []
        p["is_recommended"] = (p == min_p)

        if p["type"] == "rental" and p["rental_cost"] < p["taxi_cost"]:
            diff = p["taxi_cost"] - p["rental_cost"]
            reasons.append(f"🚗 レンタカー推奨（タクシーより {diff:,} 円お得）")
        elif p["type"] == "taxi" and p["taxi_cost"] <= p["rental_cost"]:
            reasons.append("🚕 タクシー利用推奨（費用削減）")

        if p == min_p:
            reasons.insert(0, "🏆 全ルート中最安")

        p["recommend_reason"] = " / ".join(reasons)

    return patterns, has_flight


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗", layout="wide")
st.title("🚗 交通費・出張見積もりアプリ")

# --- サイドバー ---
st.sidebar.header("🔑 APIキー設定")
gemini_api_key = st.sidebar.text_input("Gemini API Key", type="password")
navitime_api_key = st.sidebar.text_input("NAVITIME API Key (RapidAPI)", type="password")

if not gemini_api_key or not navitime_api_key:
    st.info("👈 サイドバーから Gemini API Key と NAVITIME API Key を入力してください。")
    st.stop()

# --- 入力フォーム ---
col1, col2 = st.columns([2, 1])

with col1:
    address_input = st.text_input(
        "目的地（住所または施設名）",
        "〒891-9213 鹿児島県大島郡知名町瀬利覚2208（沖永良部徳洲会病院）",
    )

with col2:
    station_choice = st.selectbox(
        "出発拠点",
        options=["淀屋橋駅", "大宮駅", "両方試算して比較"],
        index=0,
    )

col_a, col_b = st.columns(2)
with col_a:
    headcount = st.number_input("作業人数（人）", min_value=1, value=1)
with col_b:
    work_days = st.number_input("現地作業日数（日）", min_value=1, value=2)

st.markdown("---")

# --- 試算実行 ---
if st.button("🚀 最速出張見積もりを計算する（NAVITIME判定）", type="primary"):

    # 1. Gemini住所正規化
    with st.spinner("🤖 Geminiで住所を正規化中..."):
        ai_info = analyze_and_normalize_with_gemini(address_input, gemini_api_key)

    normalized_address = ai_info.get("normalized_address", address_input)
    st.success(f"**正規化後の住所:** {normalized_address}")

    # 2. ジオコーディング
    with st.spinner("📍 国土地理院APIで目的地の位置情報を取得中..."):
        geo = geocode_address(normalized_address)

    if geo is None:
        st.error("❌ 住所のジオコーディングに失敗しました。正しい住所を入力してください。")
        st.stop()

    dest_lat, dest_lon = geo

    # 3. 試算と比較
    stations_to_process = (
        STATION_COORDS if station_choice == "両方試算して比較" else {station_choice: STATION_COORDS[station_choice]}
    )

    for st_name, (st_lat, st_lon) in stations_to_process.items():
        st.markdown(f"### 🚉 出発地: 【{st_name}】")

        with st.spinner(f"🧭 {st_name} からの最速ルートをNAVITIME APIで全検索中..."):
            fastest_route = get_navitime_fastest_route(st_lat, st_lon, dest_lat, dest_lon, navitime_api_key)

        if not fastest_route:
            st.error(f"❌ {st_name} からのルート情報がNAVITIME APIから取得できませんでした。")
            continue

        patterns, has_flight = build_fastest_route_patterns(
            st_name, st_lat, st_lon, dest_lat, dest_lon, fastest_route, headcount, work_days
        )

        transport_label = f"✈️ 飛行機ルート（到着: {fastest_route['end_station']}）" if has_flight else f"🚄 新幹線/電車ルート（到着: {fastest_route['end_station']}）"
        st.info(f"💡 **【NAVITIME 最速判定結果】** この目的地への最速交通手段は **{transport_label}** （片道約 {fastest_route['time_min']} 分）です。")

        for p in patterns:
            tag_text = f"⭐ {p['recommend_reason']}" if p["recommend_reason"] else ""
            
            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円  {tag_text}", expanded=p["is_recommended"]):
                if p["is_recommended"]:
                    st.success(f"推奨パターン: {p['recommend_reason']}")

                st.write(f"🚉 **到着駅/空港:** {fastest_route['end_station']}")
                st.write(f"⏱ **片道所要時間（NAVITIME算出）:** 約 {p['time_min']} 分")
                st.write(f"🏨 **出張宿泊条件:** {p['note']}")
                st.write("**💰 費目別内訳（1,000円単位切り上げ済み）:**")
                
                for item, amt in p["breakdown"].items():
                    st.write(f"　・ {item}: **{amt:,}** 円")

                # Excel出力ボタン
                excel_bytes = create_excel_report(p, normalized_address, headcount, work_days, st_name, fastest_route['end_station'])
                st.download_button(
                    label=f"📥 この最速試算内容でExcel見積書をダウンロード (.xlsx)",
                    data=excel_bytes,
                    file_name=f"交通費見積書_{st_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{st_name}_{p['type']}"
                )

        st.markdown("---")


