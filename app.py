# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ

構成:
  入力(住所・日数) 
  → Gemini (住所正規化 & 飛行機必要性判断)
  → 国土地理院 (無料ジオコーディング)
  → NAVITIME API (公共交通・運賃取得) + フォールバック機能
  → パターン比較・1,000円単位切り上げ計算 → UI表示

必要ライブラリ:
  pip install streamlit google-generativeai requests
"""

import json
import math
import time
import requests
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

# 利用モデルの設定
GEMINI_MODEL_NAME = "gemini-2.5-flash"
GEMINI_FALLBACK_MODEL = "gemini-1.5-flash"

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


# ============================================================
# STEP 1: Gemini による住所正規化 & 出張条件の事前分析
# ============================================================

def analyze_and_normalize_with_gemini(raw_address: str, api_key: str) -> dict:
    """
    住所の正規化に加え、離島・遠方フラグや飛行機利用の必要性をAIで総合判定する。
    """
    genai.configure(api_key=api_key)

    prompt = f"""
以下の目的地の住所情報を確認し、JSON形式で回答してください。

住所入力: {raw_address}

【判定項目】
1. normalized_address: 正式な日本の住所表記（都道府県〜番地まで。建物名があれば含める）
2. requires_flight: 離島（沖永良部島、奄美大島、沖縄等）や羽田/伊丹/関空等の飛行機移動が一般的な地域の場合 true、それ以外は false
3. nearest_airport_name: 飛行機が必要な場合、目的地の最寄り空港名（例: 沖永良部空港）。不要なら null
4. flight_estimate_fare_one_way: 飛行機が必要な場合の目安片道航空券代（ANA/JALフレックス大人想定円。例: 62000）。不要なら 0

【出力フォーマット（JSONのみ、思考プロセス不可）】
{{
  "normalized_address": "鹿児島県大島郡知名町瀬利覚2208",
  "requires_flight": true,
  "nearest_airport_name": "沖永良部空港",
  "flight_estimate_fare_one_way": 62000
}}
"""

    try:
        model = genai.GenerativeModel(GEMINI_MODEL_NAME)
        res = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"},
        )
        return json.loads(res.text)
    except Exception:
        # フォールバックモデル
        try:
            model = genai.GenerativeModel(GEMINI_FALLBACK_MODEL)
            res = model.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"},
            )
            return json.loads(res.text)
        except Exception as e:
            st.warning(f"Gemini解析警告: {e}。基本住所処理に移行します。")
            return {
                "normalized_address": raw_address,
                "requires_flight": "沖永良部" in raw_address or "沖縄" in raw_address or "奄美" in raw_address,
                "nearest_airport_name": "地方空港" if "沖永良部" in raw_address else None,
                "flight_estimate_fare_one_way": 60000 if "沖永良部" in raw_address else 0,
            }


# ============================================================
# STEP 2: 国土地理院APIでジオコーディング（無料）
# ============================================================

def geocode_address(address: str):
    """住所文字列 → (lat, lon)。取得失敗時は None。"""
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
    NAVITIME Route (totalnavi) APIで2点間の公共交通経路を取得する。
    """
    headers = {
        "X-RapidAPI-Key": api_key.strip(),
        "X-RapidAPI-Host": NAVITIME_HOST,
    }
    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "format": "json",
    }

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
    except requests.exceptions.RequestException as e:
        st.error(f"❌ NAVITIME 通信エラー: {e}")
        return None

    # HTTPステータスコード別のエラーハンドリング
    if res.status_code == 401:
        st.error("❌ NAVITIME API エラー (401 Unauthorized): RapidAPIキーが無効です。")
        return None
    elif res.status_code == 403:
        st.error(
            "❌ NAVITIME API エラー (403 Forbidden): RapidAPI上で『NAVITIME Route(totalnavi)』の"
            "プランに契約（Subscribe）されていない可能性があります。\n"
            "RapidAPI画面で『Subscribe to Test / Basic Plan』ボタンを押して有効化してください。"
        )
        try:
            st.json(res.json())
        except Exception:
            st.text(res.text)
        return None
    elif res.status_code == 429:
        st.error("❌ NAVITIME API エラー (429 Rate Limit): リクエスト回数制限に達しました。")
        return None
    elif res.status_code != 200:
        st.error(f"❌ NAVITIME API エラー ({res.status_code}): {res.text}")
        return None

    data = res.json()

    # レスポンス構造から所要時間・運賃を抽出
    try:
        item = data["items"][0]
        summary = item["summary"]
        time_min = summary["move"]["time"]
        
        # 運賃の抽出（fareオブジェクト内のunit_0等）
        fare = 0
        if "fare" in summary["move"]:
            fare_info = summary["move"]["fare"]
            if isinstance(fare_info, dict):
                fare = fare_info.get("unit_0", 0) or fare_info.get("total", 0)

        # 徒歩時間算出
        walk_time_min = 0
        for section in item.get("sections", []):
            if section.get("type") == "move" and section.get("move") == "walk":
                walk_time_min += section.get("time", 0)

        return {
            "time_min": time_min,
            "fare": int(fare),
            "walk_time_min": walk_time_min,
            "raw_data": data,
        }
    except (KeyError, IndexError):
        st.warning("⚠️ NAVITIMEのレスポンスから経路情報を抽出できませんでした。データ構造をご確認ください。")
        with st.expander("レスポンスJSONを表示"):
            st.json(data)
        return None


# ============================================================
# パターン別費用計算ロジック
# ============================================================

def build_patterns(
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
    社内ルールに基づいて、各出張パターンの概算見積もりを計算する。
    """
    patterns = []

    # 前泊・後泊・宿泊日数の決定
    # 飛行機または移動4時間以上の場合は移動専用日を考慮
    travel_time_hours = (route_info["time_min"] / 60.0) if route_info else 5.0
    is_long_distance = travel_time_hours >= 4.0 or ai_info.get("requires_flight", False)

    if is_long_distance:
        # 前泊 + 後泊などで宿泊数増加想定（作業日数 + 1泊）
        nights = work_days + 1
        stay_note = f"移動時間（約{travel_time_hours:.1f}h）または飛行機利用のため、前泊/後泊想定（{nights}泊）"
    elif travel_time_hours >= 2.5:
        nights = work_days
        stay_note = f"移動2.5時間以上のため宿泊想定（{nights}泊）"
    else:
        nights = max(work_days - 1, 0)
        stay_note = f"日帰り可能または作業日数に応じた宿泊（{nights}泊）"

    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)

    # 直線距離の算出（最寄り駅・現地タクシー計算用）
    dist_km = haversine_km(station_lat, station_lon, dest_lat, dest_lon)

    # ----------------------------------------------------
    # パターン1: 新幹線/電車 + 徒歩
    # ----------------------------------------------------
    if route_info and route_info["walk_time_min"] < TAXI_WALK_THRESHOLD_MIN:
        train_fare_rt = round_up_1000(route_info["fare"] * 2 * headcount)
        total_cost = train_fare_rt + hotel_cost
        patterns.append({
            "name": "新幹線/電車 + 徒歩",
            "time_min": route_info["time_min"],
            "cost": total_cost,
            "breakdown": {
                "電車・新幹線運賃(往復・人数分)": train_fare_rt,
                "宿泊費": hotel_cost,
            },
            "note": stay_note,
        })

    # ----------------------------------------------------
    # パターン2: 新幹線/電車 + 現地タクシー
    # ----------------------------------------------------
    if route_info:
        train_fare_rt = round_up_1000(route_info["fare"] * 2 * headcount)
        
        # 現地タクシー運賃試算（徒歩時間から距離推定）
        est_taxi_km = max(route_info["walk_time_min"] * 0.08, 2.0)
        taxi_one_way = est_taxi_km * TAXI_FARE_PER_KM
        taxi_trips = nights + 1  # 滞在中の往復移動回数
        taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

        total_cost = train_fare_rt + taxi_total + hotel_cost
        patterns.append({
            "name": "新幹線/電車 + 現地タクシー",
            "time_min": route_info["time_min"] - route_info["walk_time_min"] + 15,
            "cost": total_cost,
            "breakdown": {
                "電車・新幹線運賃(往復・人数分)": train_fare_rt,
                f"現地タクシー運賃(往復×{taxi_trips}回分)": taxi_total,
                "宿泊費": hotel_cost,
            },
            "note": stay_note,
        })

    # ----------------------------------------------------
    # パターン3: 飛行機(ANA等) + 空港アクセス電車 + 現地タクシー
    # ----------------------------------------------------
    if ai_info.get("requires_flight", False) or dist_km > 300:
        flight_one_way = ai_info.get("flight_estimate_fare_one_way") or 50000
        flight_rt_total = round_up_1000(flight_one_way * 2 * headcount)

        # 空港までのアクセス電車代概算（伊丹/羽田等想定）
        airport_train_rt = round_up_1000(2000 * 2 * headcount)

        # 目的地でのタクシー代概算（空港〜目的地の距離より算出）
        # 離島や空港からの移動は20km〜30km前後と仮定
        est_airport_taxi_km = 25.0
        taxi_one_way = est_airport_taxi_km * TAXI_FARE_PER_KM
        taxi_trips = nights + 1
        taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

        total_cost = flight_rt_total + airport_train_rt + taxi_total + hotel_cost
        airport_name = ai_info.get("nearest_airport_name") or "最寄り空港"

        patterns.append({
            "name": f"飛行機(ANA/JAL) + 空港タクシー [{airport_name}経由]",
            "time_min": 300,  # 概算5時間
            "cost": total_cost,
            "breakdown": {
                "航空券費用(往復・人数分)": flight_rt_total,
                "最寄り空港アクセス電車運賃": airport_train_rt,
                f"現地タクシー運賃(空港↔目的地 {taxi_trips}往復)": taxi_total,
                "宿泊費": hotel_cost,
            },
            "note": f"飛行機必須ルート。{stay_note}",
        })

    return patterns


# ============================================================
# Streamlit UIメイン画面
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗", layout="wide")
st.title("🚗 交通費・出張見積もりアプリ")

# --- サイドバー設定 ---
st.sidebar.header("🔑 APIキー設定")
gemini_api_key = st.sidebar.text_input("Gemini API Key (AI Studio)", type="password")
navitime_api_key = st.sidebar.text_input("NAVITIME API Key (RapidAPI)", type="password")

st.sidebar.markdown("---")
st.sidebar.caption("💡 地理院ジオコーディングは完全無料で自動利用されます。")

# 接続テスト用ボタン
if st.sidebar.button("🔍 NAVITIME API 接続テスト"):
    if not navitime_api_key:
        st.sidebar.error("NAVITIME API Key を入力してください。")
    else:
        with st.sidebar.spinner("通信テスト中..."):
            test_res = get_navitime_route(
                STATION_COORDS["淀屋橋駅"][0],
                STATION_COORDS["淀屋橋駅"][1],
                34.702485,
                135.495951,
                navitime_api_key,
            )
            if test_res:
                st.sidebar.success("✅ NAVITIME API 接続成功！")

if not gemini_api_key or not navitime_api_key:
    st.info("👈 左側のサイドバーから Gemini API Key と NAVITIME API Key を入力してください。")
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
if st.button("🚀 出張見積もりを計算する", type="primary"):

    # 1. Geminiによる住所・飛行機判定
    with st.spinner("🤖 1. Geminiで住所と出張条件を分析中..."):
        ai_info = analyze_and_normalize_with_gemini(address_input, gemini_api_key)

    normalized_address = ai_info.get("normalized_address", address_input)

    st.success(f"**正規化後の住所:** {normalized_address}")
    if ai_info.get("requires_flight"):
        st.warning(f"✈️ 飛行機移動が必要なエリア（最寄り: {ai_info.get('nearest_airport_name', '空港')}）と判定されました。")

    # 2. ジオコーディング
    with st.spinner("📍 2. 国土地理院APIで目的地の緯度経度を取得中..."):
        geo = geocode_address(normalized_address)

    if geo is None:
        st.error("❌ 住所のジオコーディングに失敗しました。正式な住所表記を入力してください。")
        st.stop()

    dest_lat, dest_lon = geo
    st.info(f"📍 目的地の座標: 緯度 {dest_lat:.6f} / 経度 {dest_lon:.6f}")

    # 3. 拠点ごとのルート取得 & パターン計算
    stations_to_process = (
        STATION_COORDS if station_choice == "両方試算して比較" else {station_choice: STATION_COORDS[station_choice]}
    )

    st.markdown("### 📊 試算結果・比較一覧")

    for st_name, (st_lat, st_lon) in stations_to_process.items():
        st.markdown(f"#### 🚉 出発地: 【{st_name}】")

        with st.spinner(f"🧭 {st_name} からの最寄ルート・運賃を計算中..."):
            route_info = get_navitime_route(st_lat, st_lon, dest_lat, dest_lon, navitime_api_key)

        # パターン構築
        patterns = build_patterns(
            st_name, st_lat, st_lon, dest_lat, dest_lon, route_info, ai_info, headcount, work_days
        )

        if not patterns:
            st.error("有効な試算パターンが見つかりませんでした。")
            continue

        # 最安・最速判定
        cheapest_p = min(patterns, key=lambda x: x["cost"])
        fastest_p = min(patterns, key=lambda x: x["time_min"])

        for p in patterns:
            is_cheapest = (p == cheapest_p)
            
            # 見出しタグ
            tags = []
            if is_cheapest:
                tags.append("💰 最安（推奨）")
            if p == fastest_p and not is_cheapest:
                tags.append("⏱ 最速")

            tag_str = " ".join(tags)

            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円  {tag_str}"):
                st.write(f"**⏱ 片道所要時間（目安）:** 約 {p['time_min']} 分")
                st.write(f"**🏨 備考/条件:** {p['note']}")
                st.write("**💰 費目別内訳（1,000円単位切り上げ済み）:**")
                
                for item, amt in p["breakdown"].items():
                    st.write(f"　・ {item}: **{amt:,}** 円")

        st.markdown("---")

