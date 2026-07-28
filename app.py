# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ

機能:
  - 目的地を入力し、新幹線/電車ルートと飛行機ルートの所要時間・運賃を比較
  - 飛行機の方が速い場合は自動的に「飛行機推奨」判定を行う
  - 費目は社内ルールに従い1,000円単位で切り上げ
  - NAVITIME API (RapidAPI) + Google Gemini API 連携
  - 出発日時(start_time)の自動生成によりNAVITIME 400エラーを解決

必要ライブラリ:
  pip install streamlit google-generativeai requests
"""

import json
import math
import time
import requests
from datetime import datetime, timedelta
import streamlit as st
import google.generativeai as genai

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
    """
    お使いのAPIキーで現在呼び出し可能なGeminiモデル名を自動取得する関数。
    """
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
# STEP 1: Gemini による住所正規化 & 飛行機必要性のアシスト
# ============================================================

def analyze_and_normalize_with_gemini(raw_address: str, api_key: str) -> dict:
    """
    住所の正規化と最寄り空港・航空券価格相場の調査をGeminiで行う。
    """
    genai.configure(api_key=api_key.strip())

    prompt = f"""
以下の目的地の住所情報を確認し、JSON形式で回答してください。

住所入力: {raw_address}

【判定項目】
1. normalized_address: 正式な日本の住所表記（都道府県〜番地まで）
2. nearest_airport_name: 目的地の最寄り空港名（例: 沖永良部空港、鹿児島空港、伊丹空港など）
3. flight_estimate_fare_one_way: 普通運賃（ANA/JALフレックス）の片道大人概算運賃（円）
4. flight_duration_min: 最寄り拠点（羽田/伊丹等）からの飛行時間＋手続き等の総所要目安（分）

【出力フォーマット（JSONのみ）】
{{
  "normalized_address": "鹿児島県大島郡知名町瀬利覚2208",
  "nearest_airport_name": "沖永良部空港",
  "flight_estimate_fare_one_way": 62000,
  "flight_duration_min": 240
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
        is_island = "沖永良部" in raw_address or "沖縄" in raw_address or "奄美" in raw_address
        return {
            "normalized_address": raw_address,
            "nearest_airport_name": "最寄り空港" if is_island else "伊丹空港/羽田空港",
            "flight_estimate_fare_one_way": 60000 if is_island else 35000,
            "flight_duration_min": 240 if is_island else 180,
        }


# ============================================================
# STEP 2: 国土地理院APIでジオコーディング（無料）
# ============================================================

def geocode_address(address: str):
    """住所文字列 → (lat, lon)"""
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
# STEP 3: NAVITIME API で経路・運賃・所要時間を取得
# ============================================================

def get_navitime_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    """
    NAVITIME Route (totalnavi) APIで公共交通機関（陸路）の経路を取得する。
    """
    clean_key = api_key.strip()
    headers = {
        "X-RapidAPI-Key": clean_key,
        "X-RapidAPI-Host": NAVITIME_HOST,
    }

    # NAVITIME API必須パラメータ「start_time（出発日時）」を生成（例: 翌日の朝9時）
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

    try:
        item = data["items"][0]
        summary = item["summary"]
        time_min = summary["move"]["time"]
        
        fare = 0
        if "fare" in summary["move"]:
            fare_info = summary["move"]["fare"]
            if isinstance(fare_info, dict):
                fare = fare_info.get("unit_0", 0) or fare_info.get("total", 0)

        walk_time_min = 0
        for section in item.get("sections", []):
            if section.get("type") == "move" and section.get("move") == "walk":
                walk_time_min += section.get("time", 0)

        return {
            "time_min": time_min,
            "fare": int(fare),
            "walk_time_min": walk_time_min,
        }
    except (KeyError, IndexError):
        st.warning("⚠️ レスポンス形式の解析で予期せぬデータ構造が検出されました。")
        with st.expander("受け取ったデータ"):
            st.json(data)
        return None


# ============================================================
# 時間比較 & パターン別費用計算ロジック
# ============================================================

def build_patterns_with_time_comparison(
    station_name: str,
    station_lat: float,
    station_lon: float,
    dest_lat: float,
    dest_lon: float,
    route_info: dict,
    ai_info: dict,
    headcount: int,
    work_days: int,
):
    """
    新幹線ルートと飛行機ルートの所要時間を比較し、速い方を推奨フラグ立てする。
    """
    patterns = []
    
    # 直線距離の算出
    dist_km = haversine_km(station_lat, station_lon, dest_lat, dest_lon)

    # 電車/新幹線ルートの所要時間（取得できている場合）
    train_time_min = route_info["time_min"] if route_info else 9999

    # 飛行機ルートの総所要時間見積もり（空港移動・搭乗手続き60分含む）
    flight_time_min = ai_info.get("flight_duration_min", 240) + 60

    # ----------------------------------------------------
    # 1. 新幹線／電車ルートの組み立て
    # ----------------------------------------------------
    if route_info:
        travel_hours = train_time_min / 60.0
        if travel_hours >= 4.0:
            nights = work_days + 1
            stay_note = f"移動時間（約{travel_hours:.1f}時間）が長いため前泊/後泊想定（{nights}泊）"
        elif travel_hours >= 2.5:
            nights = work_days
            stay_note = f"移動2.5時間以上の為宿泊想定（{nights}泊）"
        else:
            nights = max(work_days - 1, 0)
            stay_note = f"日帰り/標準宿泊（{nights}泊）"

        hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)
        train_fare_rt = round_up_1000(route_info["fare"] * 2 * headcount)

        # タクシー利用計算
        est_taxi_km = max(route_info["walk_time_min"] * 0.08, 2.0)
        taxi_one_way = est_taxi_km * TAXI_FARE_PER_KM
        taxi_trips = nights + 1
        taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

        total_cost = train_fare_rt + taxi_total + hotel_cost

        patterns.append({
            "type": "train",
            "name": "新幹線/電車 + 現地タクシー",
            "time_min": train_time_min,
            "cost": total_cost,
            "breakdown": {
                "電車・新幹線運賃(往復・人数分)": train_fare_rt,
                f"現地タクシー運賃(往復×{taxi_trips}回分)": taxi_total,
                "宿泊費": hotel_cost,
            },
            "note": stay_note,
        })

    # ----------------------------------------------------
    # 2. 飛行機ルートの組み立て
    # ----------------------------------------------------
    is_island = "沖永良部" in ai_info.get("normalized_address", "") or "沖縄" in ai_info.get("normalized_address", "")
    if is_island or dist_km > 250 or flight_time_min < train_time_min:
        
        nights = work_days + 1
        stay_note = f"飛行機移動のため前泊/後泊想定（{nights}泊）"
        hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)

        flight_one_way = ai_info.get("flight_estimate_fare_one_way", 50000)
        flight_rt_total = round_up_1000(flight_one_way * 2 * headcount)
        airport_access_train = round_up_1000(2000 * 2 * headcount)

        est_airport_taxi_km = 20.0 if not is_island else 15.0
        taxi_one_way = est_airport_taxi_km * TAXI_FARE_PER_KM
        taxi_trips = nights + 1
        taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

        total_cost = flight_rt_total + airport_access_train + taxi_total + hotel_cost
        airport_name = ai_info.get("nearest_airport_name", "最寄り空港")

        patterns.append({
            "type": "flight",
            "name": f"飛行機(ANA/JAL) + 空港アクセス [{airport_name}利用]",
            "time_min": flight_time_min,
            "cost": total_cost,
            "breakdown": {
                "航空券費用(往復・人数分)": flight_rt_total,
                "最寄り空港アクセス電車運賃": airport_access_train,
                f"現地タクシー運賃(空港↔目的地 {taxi_trips}往復)": taxi_total,
                "宿泊費": hotel_cost,
            },
            "note": stay_note,
        })

    # ----------------------------------------------------
    # 3. 飛行機 vs 新幹線 所要時間チェック & 判定
    # ----------------------------------------------------
    train_pattern = next((p for p in patterns if p["type"] == "train"), None)
    flight_pattern = next((p for p in patterns if p["type"] == "flight"), None)

    is_flight_faster = False
    time_diff_min = 0

    if train_pattern and flight_pattern:
        if flight_pattern["time_min"] < train_pattern["time_min"]:
            is_flight_faster = True
            time_diff_min = train_pattern["time_min"] - flight_pattern["time_min"]

    for p in patterns:
        p["is_recommended"] = False
        p["recommend_reason"] = ""

        if p["type"] == "flight" and (is_flight_faster or is_island):
            p["is_recommended"] = True
            if is_flight_faster:
                p["recommend_reason"] = f"✈️ 飛行機推奨（新幹線より約 {time_diff_min} 分短縮可能）"
            else:
                p["recommend_reason"] = "✈️ 飛行機推奨（離島・遠方ルート）"
        elif p["type"] == "train" and not is_flight_faster and not is_island:
            p["is_recommended"] = True
            p["recommend_reason"] = "🚄 新幹線/電車推奨（陸路アクセス推奨）"

    return patterns, is_flight_faster, time_diff_min


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
if st.button("🚀 見積もり・最速ルートを計算する", type="primary"):

    # 1. Geminiで住所正規化・飛行機見積情報取得
    with st.spinner("🤖 Geminiで目的地と移動手段を分析中..."):
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

    # 3. 試算と時間比較
    stations_to_process = (
        STATION_COORDS if station_choice == "両方試算して比較" else {station_choice: STATION_COORDS[station_choice]}
    )

    for st_name, (st_lat, st_lon) in stations_to_process.items():
        st.markdown(f"### 🚉 出発地: 【{st_name}】")

        with st.spinner(f"🧭 {st_name} からの経路を計算中..."):
            route_info = get_navitime_route(st_lat, st_lon, dest_lat, dest_lon, navitime_api_key)

        patterns, is_flight_faster, time_diff = build_patterns_with_time_comparison(
            st_name, st_lat, st_lon, dest_lat, dest_lon, route_info, ai_info, headcount, work_days
        )

        if is_flight_faster:
            st.success(f"💡 **【速度判定】** 飛行機を利用した方が新幹線/電車より **約 {time_diff} 分短縮** できます！飛行機ルートを優先選定しました。")

        for p in patterns:
            tag_text = f"⭐ {p['recommend_reason']}" if p["is_recommended"] else ""
            
            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円  {tag_text}", expanded=p["is_recommended"]):
                if p["is_recommended"]:
                    st.success(p["recommend_reason"])

                st.write(f"⏱ **片道所要時間（目安）:** 約 {p['time_min']} 分")
                st.write(f"🏨 **出張宿泊条件:** {p['note']}")
                st.write("**💰 費目別内訳（1,000円単位切り上げ済み）:**")
                
                for item, amt in p["breakdown"].items():
                    st.write(f"　・ {item}: **{amt:,}** 円")

        st.markdown("---")

