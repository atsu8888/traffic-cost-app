# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（プロトタイプ）

構成:
  入力(住所) → Gemini(住所正規化) → GSI(ジオコーディング) → NAVITIME(経路・運賃)
  → ルールベースでパターン比較・料金計算 → 画面表示

前提:
  - 出発地は「大宮駅」「淀屋橋駅」の2拠点固定（座標は下記 STATION_COORDS に直書き）
  - 目的地は住所ベースの自由入力を想定
  - 飛行機(ANA)パターンとタクシーの実費計算は、外部の運賃データ/APIが
    別途必要なため、このプロトタイプでは「仮の単価」でスタブ実装しています。
    本番導入時は TODO の箇所を実データに差し替えてください。
"""

import json
import math
import requests
import streamlit as st
import google.generativeai as genai

# ============================================================
# 定数・固定値（社内ルールより）
# ============================================================

STATION_COORDS = {
    "大宮駅": (35.906702, 139.623596),
    "淀屋橋駅": (34.693729, 135.499814),
}

HOTEL_COST_PER_NIGHT_PER_PERSON = 20_000   # 宿泊費（1泊/1名）
RENTAL_CAR_COST_PER_DAY = 12_000           # レンタカー（1日）
TAXI_WALK_THRESHOLD_MIN = 15               # 最寄駅から徒歩何分以上でタクシー検討するか

# TODO: 実際のタクシー相場に合わせて調整、または実費APIに置き換える
TAXI_FARE_PER_KM = 400  # 円/km（初乗り+距離想定の簡易単価。地域により要調整）

GEMINI_MODEL_NAME = "models/gemini-flash-latest"
NAVITIME_URL = "https://navitime-route-totalnav.p.rapidapi.com/route_transit"
NAVITIME_HOST = "navitime-route-totalnav.p.rapidapi.com"
GSI_GEOCODE_URL = "https://msearch.gsi.go.jp/address-search/AddressSearch"


# ============================================================
# STEP 1: Gemini による住所正規化
# ============================================================

def normalize_address_with_gemini(raw_address: str, api_key: str) -> dict:
    """
    自由入力の住所を、日本の正式な住所表記に正規化する。
    座標計算はさせない（ハルシネーション対策）。あくまで表記ゆれの補完のみ。
    """
    genai.configure(api_key=api_key)

    prompt = f"""
以下の入力を、日本の正式な住所表記（都道府県から番地まで）に正規化してください。
郵便番号や建物名が含まれる場合はそのまま活用し、省略や誤字があれば補完してください。
座標や緯度経度は絶対に含めないでください。住所の文字列のみを扱ってください。

入力: {raw_address}

以下のJSON形式のみで返答してください。前置きや説明文は一切不要です。
{{"normalized_address": "広島県広島市中区中島町3番30号"}}
"""

    model = genai.GenerativeModel(GEMINI_MODEL_NAME)
    res = model.generate_content(
        prompt,
        generation_config={"response_mime_type": "application/json"},
    )
    data = json.loads(res.text)
    return data


# ============================================================
# STEP 2: 国土地理院APIでジオコーディング（無料・キー不要）
# ============================================================

def geocode_address(address: str):
    """住所文字列 → (lat, lon)。見つからない場合は None。"""
    try:
        res = requests.get(GSI_GEOCODE_URL, params={"q": address}, timeout=10)
        res.raise_for_status()
        data = res.json()
    except requests.exceptions.RequestException:
        return None

    if not data:
        return None

    lon, lat = data[0]["geometry"]["coordinates"]  # GSIは[経度, 緯度]の順
    return lat, lon


# ============================================================
# STEP 3: NAVITIME API で経路・運賃・所要時間を取得
# ============================================================

def get_navitime_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    """
    2点間の公共交通ルートを取得する。
    戻り値: {"time_min": int, "fare": int, "walk_time_min": int} または None
    """
    headers = {
        "X-RapidAPI-Key": api_key,
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
        st.error(f"NAVITIME 通信エラー: {e}")
        return None

    if res.status_code != 200:
        st.error(f"NAVITIME API エラー: {res.status_code}")
        return None

    data = res.json()

    # NAVITIME のレスポンス構造からサマリーを抽出
    # 実際のレスポンス構造に応じて調整が必要（items[0].summary.move など）
    try:
        item = data["items"][0]
        summary = item["summary"]
        time_min = summary["move"]["time"]
        fare = summary["move"].get("fare", {}).get("unit_0", 0)

        # 徒歩区間の合計時間を section から抽出（存在する場合）
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
        st.warning("NAVITIMEのレスポンス形式が想定と異なります。レスポンス全体を確認してください。")
        st.json(data)
        return None


# ============================================================
# 距離計算（タクシー概算用・簡易straight-line距離）
# ============================================================

def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    )
    return 2 * r * math.asin(math.sqrt(a))


# ============================================================
# 料金計算ロジック（社内ルールをそのままコード化）
# ============================================================

def round_up_1000(amount: float) -> int:
    """各費目を1,000円単位で切り上げ"""
    return int(math.ceil(amount / 1000.0) * 1000)


def build_patterns(route_info: dict, dest_lat, dest_lon, headcount: int, nights: int):
    """
    経路情報とルールから、パターンごとの費用・所要時間を組み立てる。
    現状は「電車+徒歩」「電車+タクシー」の2パターンのみ実装。
    飛行機(ANA)パターンは運賃データが必要なため TODO。
    """
    patterns = []

    train_fare_roundtrip = route_info["fare"] * 2 * headcount
    train_fare_roundtrip = round_up_1000(train_fare_roundtrip)
    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)

    # --- パターン1: 電車 + 徒歩 ---
    if route_info["walk_time_min"] < TAXI_WALK_THRESHOLD_MIN:
        total_time = route_info["time_min"]
        total_cost = train_fare_roundtrip + hotel_cost
        patterns.append({
            "name": "新幹線/電車 + 徒歩",
            "time_min": total_time,
            "cost": total_cost,
            "breakdown": {
                "電車運賃(往復・人数分)": train_fare_roundtrip,
                "宿泊費": hotel_cost,
            },
        })

    # --- パターン2: 電車 + タクシー(現地) ---
    # タクシー: 最寄駅~目的地の距離を簡易直線距離で概算(TODO: 実際の道路距離API推奨)
    taxi_distance_km = haversine_km(dest_lat, dest_lon, dest_lat, dest_lon)  # placeholder 0
    # 実際には「NAVITIMEの最寄駅座標」からの距離を使うべきだが、
    # 現状のNAVITIMEレスポンス抽出では最寄駅座標を取得していないため、
    # 徒歩時間から概算距離を逆算する簡易ロジックにしている(TODO: 要精緻化)
    estimated_distance_km = max(route_info["walk_time_min"] * 0.08, 0.5)  # 徒歩速度4.8km/h換算の概算
    taxi_fare_one_way = estimated_distance_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1  # ルール: 宿泊日数+1日分の往復
    taxi_fare_total = round_up_1000(taxi_fare_one_way * 2 * taxi_trips)

    total_time_taxi = route_info["time_min"] - route_info["walk_time_min"] + 10  # 徒歩をタクシー移動時間(概算10分)に置換
    total_cost_taxi = train_fare_roundtrip + taxi_fare_total + hotel_cost

    patterns.append({
        "name": "新幹線/電車 + タクシー",
        "time_min": total_time_taxi,
        "cost": total_cost_taxi,
        "breakdown": {
            "電車運賃(往復・人数分)": train_fare_roundtrip,
            "タクシー運賃(往復×{}回)".format(taxi_trips): taxi_fare_total,
            "宿泊費": hotel_cost,
        },
    })

    # --- パターン3〜5: 飛行機(ANA)系 ---
    # TODO: ANA運賃データ(APIまたは固定テーブル)が必要。
    # 現時点ではダミーの運賃を使わず、明示的に「未実装」として扱う。
    if route_info["time_min"] >= 240:
        patterns.append({
            "name": "飛行機(ANA)系パターン",
            "time_min": None,
            "cost": None,
            "breakdown": {},
            "note": "移動時間が4時間以上のため本来は飛行機優先ですが、"
                    "ANA運賃データ未連携のため計算していません。運賃テーブルの追加が必要です。",
        })

    return patterns


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗")
st.title("🚗 交通費・出張見積もりアプリ")

st.sidebar.header("🔑 APIキー設定")
gemini_api_key = st.sidebar.text_input("Gemini API Key (AI Studio)", type="password")
navitime_api_key = st.sidebar.text_input("NAVITIME API Key (RapidAPI)", type="password")

st.sidebar.caption("※ 国土地理院のジオコーディングAPIはキー不要です")

if not gemini_api_key or not navitime_api_key:
    st.warning("👈 左側のサイドバーからAPIキーを入力してください。")
    st.stop()

address_input = st.text_input("目的地住所", "広島県広島市中区中島町3番30号")
headcount = st.number_input("人数", min_value=1, value=2)
nights = st.number_input("宿泊日数", min_value=0, value=4)

if st.button("見積もりを試算する"):

    # --- STEP 1: 住所正規化 ---
    with st.spinner("Geminiで住所を正規化中..."):
        try:
            normalized = normalize_address_with_gemini(address_input, gemini_api_key)
            normalized_address = normalized.get("normalized_address", address_input)
        except Exception as e:
            st.error(f"Gemini呼び出しでエラー: {e}")
            st.stop()

    st.subheader("① 住所正規化結果")
    st.write(normalized_address)

    # --- STEP 2: ジオコーディング ---
    with st.spinner("目的地の座標を取得中..."):
        geo = geocode_address(normalized_address)

    if geo is None:
        st.error("住所のジオコーディングに失敗しました。住所表記をご確認ください。")
        st.stop()

    dest_lat, dest_lon = geo
    st.subheader("② ジオコーディング結果")
    st.write(f"緯度: {dest_lat} / 経度: {dest_lon}")

    # --- STEP 3: 大宮駅・淀屋橋駅それぞれで経路計算 ---
    st.subheader("③ 出発駅ごとの経路・運賃")

    results = {}
    for station_name, (s_lat, s_lon) in STATION_COORDS.items():
        with st.spinner(f"{station_name}からの経路を計算中..."):
            route_info = get_navitime_route(s_lat, s_lon, dest_lat, dest_lon, navitime_api_key)

        if route_info is None:
            st.warning(f"{station_name}からの経路取得に失敗しました。")
            continue

        results[station_name] = route_info

        st.markdown(f"**{station_name} → 目的地**")
        st.write(
            f"所要時間: 約{route_info['time_min']}分 / "
            f"運賃(片道): {route_info['fare']}円 / "
            f"徒歩時間: {route_info['walk_time_min']}分"
        )

    if not results:
        st.error("経路情報が取得できませんでした。")
        st.stop()

    # --- STEP 4: パターン別費用計算 ---
    st.subheader("④ パターン別 見積もり結果")

    for station_name, route_info in results.items():
        st.markdown(f"### 出発地: {station_name}")
        patterns = build_patterns(route_info, dest_lat, dest_lon, headcount, nights)

        valid_patterns = [p for p in patterns if p["cost"] is not None]
        if valid_patterns:
            cheapest = min(valid_patterns, key=lambda p: p["cost"])
            fastest = min(valid_patterns, key=lambda p: p["time_min"])
        else:
            cheapest = fastest = None

        for p in patterns:
            with st.expander(f"{p['name']}"):
                if p["cost"] is None:
                    st.info(p.get("note", "計算未対応です。"))
                    continue

                tag = []
                if cheapest and p is cheapest:
                    tag.append("💰 費用優先(推奨)")
                if fastest and p is fastest and fastest is not cheapest:
                    tag.append("⏱ 時間優先(補足)")
                if tag:
                    st.success(" / ".join(tag))

                st.write(f"所要時間: 約{p['time_min']}分")
                st.write(f"合計費用: {p['cost']:,}円")
                st.write("内訳:")
                for label, val in p["breakdown"].items():
                    st.write(f"　- {label}: {val:,}円")
