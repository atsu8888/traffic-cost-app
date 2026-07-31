
# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（Streamlit Cloud公開版）

デプロイ方法:
  1. GitHubリポジトリに app.py と requirements.txt を配置
  2. Streamlit Cloud でデプロイ
  3. Streamlit Cloud の Secrets 設定画面で以下を登録:
     [api_keys]
     gemini = "YOUR_GEMINI_API_KEY"
     navitime = "YOUR_NAVITIME_RAPIDAPI_KEY"

必要パッケージ: requirements.txt 参照

処理フロー:
  STEP 1: Gemini（Search なし）→ 住所正規化・座標取得・離島判定
  STEP 1.5: 離島の場合のみ Gemini + Google Search → 運賃検索
  STEP 2: 本土の場合 NAVITIME → ルート検索
  STEP 3: パターン生成・比較 → 最安ルート表示

運賃計算ルール:
  - 在来線のみ → unit_0（乗車券）のみ
  - 新幹線あり → unit_0（乗車券）+ unit_3（指定席特急券）
  - 飛行機 → セクションの unit_0（flex/普通運賃）
  - グリーン車料金は常に除外
"""

import json
import math
import io
import re
import time
import unicodedata
import requests
from datetime import datetime, timedelta
import streamlit as st

# 新SDK: google-genai
from google import genai
from google.genai import types

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

HOTEL_COST_PER_NIGHT_PER_PERSON = 20_000
RENTAL_CAR_COST_PER_DAY = 12_000
TAXI_FARE_PER_KM = 400
TAXI_WALK_THRESHOLD_MIN = 15

NAVITIME_URL = "https://navitime-route-totalnavi.p.rapidapi.com/route_transit"
NAVITIME_HOST = "navitime-route-totalnavi.p.rapidapi.com"
GSI_GEOCODE_URL = "https://msearch.gsi.go.jp/address-search/AddressSearch"

FLIGHT_MOVE_TYPES = {
    "plane", "flight", "air", "airplane", "aeroplane",
    "domestic_flight", "international_flight", "domestic_air"
}

GEMINI_MODEL = "gemini-3.5-flash"

# リトライ設定
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = [5, 15, 30]
RETRYABLE_STATUS_CODES = {429, 503}

# デバッグモード（Trueにするとレスポンス詳細を表示）
DEBUG_MODE = True


# ============================================================
# APIキー取得（st.secrets から安全に読み込み）
# ============================================================

def get_api_keys() -> tuple[str, str]:
    try:
        gemini_key = st.secrets["api_keys"]["gemini"]
        navitime_key = st.secrets["api_keys"]["navitime"]
        return gemini_key, navitime_key
    except KeyError as e:
        st.error(
            "❌ APIキーが設定されていません。\n\n"
            "Streamlit Cloud の **Settings → Secrets** に以下の形式で登録してください:\n\n"
            "```toml\n"
            "[api_keys]\n"
            'gemini = "YOUR_GEMINI_API_KEY"\n'
            'navitime = "YOUR_NAVITIME_RAPIDAPI_KEY"\n'
            "```"
        )
        st.stop()
        return "", ""


# ============================================================
# Gemini API リトライラッパー
# ============================================================

def call_gemini_with_retry(client, model: str, contents: str, config, retry_status_placeholder=None) -> str:
    last_error = None

    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
            return response.text

        except Exception as e:
            last_error = e
            error_str = str(e)

            is_retryable = any(str(code) in error_str for code in RETRYABLE_STATUS_CODES)

            if not is_retryable:
                raise e

            if attempt >= MAX_RETRIES - 1:
                break

            wait_sec = RETRY_WAIT_SECONDS[attempt]

            if retry_status_placeholder:
                retry_status_placeholder.warning(
                    f"⏳ サーバー混雑中... {wait_sec}秒後にリトライします（{attempt + 1}/{MAX_RETRIES}回目）"
                )

            time.sleep(wait_sec)

            if retry_status_placeholder:
                retry_status_placeholder.empty()

    raise last_error


# ============================================================
# ユーティリティ関数
# ============================================================

def round_up_1000(amount: float) -> int:
    if amount < 0:
        st.warning(f"⚠️ 負の金額が検出されました: {amount}")
        return 0
    if amount == 0:
        return 0
    return int(math.ceil(amount / 1000.0) * 1000)


def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2 +
         math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
         math.sin(dlon / 2) ** 2)
    return 2 * r * math.asin(math.sqrt(a))


def get_display_width(text: str) -> int:
    width = 0
    for char in str(text):
        if unicodedata.east_asian_width(char) in ('F', 'W', 'A'):
            width += 2
        else:
            width += 1
    return width


def guess_airport_from_address(address: str) -> str:
    mapping = [
        (["沖永良部", "知名町", "和泊町"], "沖永良部空港"),
        (["奄美", "龍郷町"], "奄美空港"),
        (["徳之島", "伊仙町", "天城町"], "徳之島空港"),
        (["与論"], "与論空港"),
        (["石垣"], "新石垣空港"),
        (["宮古"], "宮古空港"),
        (["久米島"], "久米島空港"),
        (["屋久島"], "屋久島空港"),
        (["種子島"], "種子島空港"),
        (["喜界"], "喜界空港"),
        (["沖縄", "那覇", "宜野湾", "浦添"], "那覇空港"),
    ]
    for keywords, airport in mapping:
        if any(kw in address for kw in keywords):
            return airport
    return "最寄り空港"


def parse_json_from_text(text: str) -> dict:
    start_idx = text.find('{')
    if start_idx == -1:
        return {}

    depth = 0
    end_idx = start_idx
    for i in range(start_idx, len(text)):
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                end_idx = i
                break

    if depth != 0:
        match = re.search(r'\{[^{}]*\}', text)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                return {}
        return {}

    json_str = text[start_idx:end_idx + 1]
    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        cleaned = re.sub(r'//.*?\n', '\n', json_str)
        cleaned = re.sub(r',\s*}', '}', cleaned)
        cleaned = re.sub(r',\s*]', ']', cleaned)
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            return {}


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
    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )
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

    travel_days = 2 if "前泊" in pattern_data.get("note", "") else 1
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

        route_val = item["route"] if is_checked else "-"
        ws.cell(row=row_idx, column=3, value=route_val).alignment = align_left

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

    for col in ws.columns:
        max_width = 0
        for cell in col:
            cell_width = get_display_width(cell.value or '')
            if cell_width > max_width:
                max_width = cell_width
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_width * 0.8 + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による目的地分析（Google Search なし・軽量版）
# ============================================================

def geocode_fallback(address: str):
    try:
        res = requests.get(GSI_GEOCODE_URL, params={"q": address}, timeout=5)
        if res.status_code == 200 and res.json():
            lon, lat = res.json()[0]["geometry"]["coordinates"]
            return lat, lon
    except Exception as e:
        st.warning(f"⚠️ 国土地理院ジオコーディング失敗: {e}")
    return None, None


def analyze_destination_with_gemini(raw_address: str, gemini_key: str, origin_name: str) -> dict:
    """
    STEP 1: Gemini AI（Google Search なし）で目的地の基本情報を分析。
    """
    client = genai.Client(api_key=gemini_key.strip())

    prompt = f"""
以下の住所・施設名について、基本情報をJSON形式で回答してください。
Google検索は不要です。あなたの知識のみで回答してください。

【対象】: {raw_address}

以下の形式のJSONテキストのみを回答してください。
{{
  "normalized_address": "対象の正式な住所（都道府県から）",
  "dest_lat": 目的地の緯度(数値),
  "dest_lon": 目的地の経度(数値),
  "is_island_or_remote": 対象が海を渡る完全な離島(沖縄本島、奄美大島、石垣島、宮古島、屋久島、種子島、徳之島、沖永良部島、与論島、久米島、喜界島など)ならtrue。北海道・本州・四国・九州の本土はfalse,
  "nearest_airport_name": "離島の場合のみ最寄り空港名(本土の場合は空文字)"
}}
"""

    islands = ["沖縄", "奄美", "沖永良部", "石垣", "宮古", "屋久島", "種子島", "徳之島", "与論", "久米島", "喜界"]
    is_island_guess = any(island in raw_address for island in islands)

    fallback_data = {
        "normalized_address": raw_address,
        "is_island_or_remote": is_island_guess,
        "nearest_airport_name": guess_airport_from_address(raw_address) if is_island_guess else "",
        "dest_lat": None,
        "dest_lon": None,
    }

    retry_placeholder = st.empty()

    try:
        config = types.GenerateContentConfig(
            temperature=0.1,
        )

        text = call_gemini_with_retry(
            client=client,
            model=GEMINI_MODEL,
            contents=prompt,
            config=config,
            retry_status_placeholder=retry_placeholder,
        )

        if text:
            parsed_data = parse_json_from_text(text)
            if parsed_data:
                fallback_data.update(parsed_data)

    except Exception as e:
        st.warning(f"⚠️ Gemini API呼び出し失敗（フォールバック値を使用）: {e}")
    finally:
        retry_placeholder.empty()

    # 座標のフォールバック
    if not fallback_data.get("dest_lat") or not fallback_data.get("dest_lon"):
        lat, lon = geocode_fallback(fallback_data["normalized_address"])
        fallback_data["dest_lat"] = lat
        fallback_data["dest_lon"] = lon

    # 空港名の補正（離島の場合）
    if fallback_data.get("is_island_or_remote"):
        guessed_airport = guess_airport_from_address(fallback_data["normalized_address"])
        if not fallback_data.get("nearest_airport_name") or fallback_data["nearest_airport_name"] == "最寄り空港":
            fallback_data["nearest_airport_name"] = guessed_airport
        elif (guessed_airport != "最寄り空港" and guessed_airport != "那覇空港"
              and fallback_data["nearest_airport_name"] == "那覇空港"):
            fallback_data["nearest_airport_name"] = guessed_airport

    return fallback_data


# ============================================================
# STEP 1.5: 離島の場合のみ Gemini + Google Search で運賃検索
# ============================================================

def search_flight_fare_with_gemini(raw_address: str, airport_name: str, gemini_key: str, origin_name: str) -> dict:
    """
    STEP 1.5: 離島の場合のみ呼ばれる。
    Gemini + Google Search Grounding でリアルタイムの運賃・所要時間を検索する。
    """
    client = genai.Client(api_key=gemini_key.strip())
    origin_city = "大阪" if "淀屋橋" in origin_name else "東京"

    prompt = f"""
出張旅費算出のためGoogle検索で調査してください。

【出発地】: {origin_city}
【到着空港】: {airport_name}
【最終目的地】: {raw_address}

以下の形式のJSONテキストのみを回答してください。
{{
  "airport_lat": 到着空港の緯度(数値),
  "airport_lon": 到着空港の経度(数値),
  "flight_fare_estimate": {origin_city}から{airport_name}への大人片道普通運賃の概算(数値・円),
  "flight_time_min": {origin_city}から{airport_name}までの片道総所要時間(数値・分)
}}
"""

    fallback_data = {
        "airport_lat": None,
        "airport_lon": None,
        "flight_fare_estimate": 60000,
        "flight_time_min": 180,
    }

    retry_placeholder = st.empty()

    try:
        google_search_tool = types.Tool(
            google_search=types.GoogleSearch()
        )

        config = types.GenerateContentConfig(
            tools=[google_search_tool],
            temperature=0.1,
        )

        text = call_gemini_with_retry(
            client=client,
            model=GEMINI_MODEL,
            contents=prompt,
            config=config,
            retry_status_placeholder=retry_placeholder,
        )

        if text:
            parsed_data = parse_json_from_text(text)
            if parsed_data:
                fallback_data.update(parsed_data)

    except Exception as e:
        st.warning(f"⚠️ 運賃検索失敗（フォールバック値を使用）: {e}")
    finally:
        retry_placeholder.empty()

    return fallback_data


# ============================================================
# STEP 2: NAVITIME API (全国・全交通手段を検索) + デバッグログ
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, navitime_key: str,
                               no_flight: bool = False, no_shinkansen: bool = False):
    """
    STEP 2: 本土の目的地に対してNAVITIME APIで最速ルートを検索。

    Args:
        no_flight: Trueの場合、飛行機を除外して検索
        no_shinkansen: Trueの場合、新幹線を除外して検索

    運賃計算ルール:
      - move_type に superexpress_train が含まれる → unit_0 + unit_3（乗車券+指定席特急券）
      - move_type に superexpress_train が含まれない → unit_0 のみ（乗車券のみ）
      - グリーン車料金は常に除外
      - 飛行機セクションの unit_0 = flex（普通運賃）として使用
    """
    clean_key = navitime_key.strip()
    headers = {"X-RapidAPI-Key": clean_key, "X-RapidAPI-Host": NAVITIME_HOST}

    future_day = datetime.now() + timedelta(days=21)
    start_time_iso = future_day.strftime("%Y-%m-%dT09:00:00")

    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "start_time": start_time_iso,
        "format": "json"
    }

    # ★ 交通手段除外パラメータ
    unuse_list = []
    if no_flight:
        unuse_list.append("domestic_flight")
    if no_shinkansen:
        unuse_list.append("superexpress_train")

    if unuse_list:
        params["unuse"] = ",".join(unuse_list)

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
        if res.status_code != 200:
            st.warning(f"⚠️ NAVITIME API エラー: HTTP {res.status_code}")
            if DEBUG_MODE:
                st.code(res.text[:2000], language="json")
            return None

        data = res.json()
        items = data.get("items", [])
        if not items:
            st.warning("⚠️ NAVITIME: ルートが見つかりませんでした。")
            if DEBUG_MODE:
                with st.expander("🔍 [DEBUG] NAVITIMEレスポンス全体"):
                    st.json(data)
            return None

        fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
        time_min = fastest_item.get("summary", {}).get("move", {}).get("time", 0)

        move_info = fastest_item.get("summary", {}).get("move", {})

        # ★ move_type を取得して新幹線の有無を判定
        move_types = move_info.get("move_type", [])
        has_superexpress = "superexpress_train" in move_types

        # ★ 運賃取得
        fare_dict = move_info.get("fare", {})
        if isinstance(fare_dict, dict):
            fare_unit_0 = fare_dict.get("unit_0", 0)  # 乗車券 + 航空券
            fare_unit_2 = fare_dict.get("unit_2", 0)  # 自由席特急券（参考表示のみ）
            fare_unit_3 = fare_dict.get("unit_3", 0)  # 指定席特急券 or グリーン車
        else:
            fare_unit_0 = 0
            fare_unit_2 = 0
            fare_unit_3 = 0

        # ★ 新幹線がある場合のみ unit_3（指定席特急券）を加算
        #    在来線のみの場合、unit_3 はグリーン車料金なので除外
        if has_superexpress:
            total_fare = int(fare_unit_0 + fare_unit_3)
            fare_note = "unit_0 + unit_3（乗車券+指定席特急券）"
        else:
            total_fare = int(fare_unit_0)
            fare_note = "unit_0のみ（乗車券）※グリーン車除外"

        has_flight = False
        flight_fare = 0

        start_station_name = None
        end_station_name = None
        end_station_lat = None
        end_station_lon = None
        flight_start_name = None
        flight_end_name = None

        sections = fastest_item.get("sections", [])

        # デバッグ: セクション詳細を収集
        debug_sections = []

        for i, sec in enumerate(sections):
            m_type = sec.get("move", "")
            sec_type = sec.get("type", "")

            if DEBUG_MODE:
                debug_sec = {
                    "index": i,
                    "type": sec_type,
                    "move": m_type,
                    "name": sec.get("name", ""),
                    "time": sec.get("time", ""),
                    "line_name": sec.get("line_name", ""),
                }
                transport = sec.get("transport", {})
                if transport and "fare" in transport:
                    debug_sec["fare"] = transport["fare"]
                debug_sections.append(debug_sec)

            if sec_type == "point":
                if not start_station_name:
                    start_station_name = sec.get("name")

                p_name = sec.get("name", "").lower()
                if "goal" not in p_name and "目的地" not in p_name:
                    end_station_name = sec.get("name")
                    if "coord" in sec:
                        end_station_lat = sec["coord"].get("lat")
                        end_station_lon = sec["coord"].get("lon")

            if sec_type == "move" and m_type.lower() in FLIGHT_MOVE_TYPES:
                has_flight = True
                # ★ 飛行機セクションの unit_0 = flex（普通運賃）
                transport = sec.get("transport", {})
                if transport and "fare" in transport and isinstance(transport["fare"], dict):
                    flight_fare += int(transport["fare"].get("unit_0", 0))

                if not flight_start_name:
                    for j in range(i - 1, -1, -1):
                        if sections[j].get("type") == "point":
                            flight_start_name = sections[j].get("name")
                            break
                for j in range(i + 1, len(sections)):
                    if sections[j].get("type") == "point":
                        flight_end_name = sections[j].get("name")
                        break

        last_walk_min = 0
        if sections and sections[-1].get("type") == "move":
            last_walk_min = sections[-1].get("time", 0)

        if has_flight:
            if flight_fare == 0:
                flight_fare = int(total_fare * 0.8)
            # アクセス電車 = 全体 - 航空券
            access_train_fare = int(total_fare - flight_fare)
        else:
            access_train_fare = 0

        result = {
            "has_flight": has_flight,
            "time_min": time_min,
            "total_fare": total_fare,
            "flight_fare": int(flight_fare),
            "access_train_fare": int(access_train_fare),
            "last_walk_min": last_walk_min,
            "start_station": start_station_name or "出発駅",
            "end_station": end_station_name or "到着駅",
            "end_station_lat": end_station_lat,
            "end_station_lon": end_station_lon,
            "flight_start": flight_start_name,
            "flight_end": flight_end_name
        }

        # ★ デバッグ表示
        if DEBUG_MODE:
            unuse_display = ", ".join(unuse_list) if unuse_list else "なし"
            with st.expander(f"🔍 [DEBUG] NAVITIMEレスポンス詳細 (除外: {unuse_display})", expanded=False):
                st.markdown("**📡 リクエストパラメータ:**")
                st.json(params)

                st.markdown("**📊 Summary (move):**")
                st.json(move_info)

                st.markdown(f"**🚄 move_type:** `{move_types}`")
                st.markdown(f"**新幹線判定:** `has_superexpress = {has_superexpress}`")

                st.markdown(f"**💰 運賃内訳 (fare):**")
                st.markdown(f"- `unit_0` (乗車券+航空券): **{fare_unit_0:,.0f}** 円")
                if has_superexpress:
                    st.markdown(f"- `unit_3` (指定席特急券): **{fare_unit_3:,.0f}** 円 ← 加算")
                else:
                    st.markdown(f"- `unit_3` (グリーン車?): ~~{fare_unit_3:,.0f}~~ 円 ← 除外（新幹線なし）")
                st.markdown(f"- `unit_2` (自由席): ~~{fare_unit_2:,.0f}~~ 円 ← 不使用")
                st.markdown(f"- **採用方式:** {fare_note}")
                st.markdown(f"- **採用合計: {total_fare:,} 円**")

                st.markdown(f"**✈️ 飛行機判定:** `has_flight = {has_flight}`")
                if has_flight:
                    st.markdown(f"- 航空券 flex (飛行機セクション unit_0): **{flight_fare:,}** 円")
                    st.markdown(f"- 電車代 (全体 - 航空券): **{access_train_fare:,}** 円")
                    st.markdown(f"- flight_start: {flight_start_name}")
                    st.markdown(f"- flight_end: {flight_end_name}")

                st.markdown("**🚏 セクション詳細:**")
                for ds in debug_sections:
                    sec_label = f"[{ds['index']}] type={ds['type']}, move={ds['move']}"
                    if ds['name']:
                        sec_label += f", name={ds['name']}"
                    if ds['line_name']:
                        sec_label += f", line={ds['line_name']}"
                    if ds['time']:
                        sec_label += f", time={ds['time']}min"
                    if ds.get('fare'):
                        f_info = ds['fare']
                        u0 = f_info.get('unit_0', '-')
                        u3 = f_info.get('unit_3', '-')
                        sec_label += f", unit_0={u0}, unit_3={u3}"
                    st.text(sec_label)

                st.markdown("**📋 最終計算結果:**")
                st.json(result)

        return result

    except requests.exceptions.Timeout:
        st.warning("⚠️ NAVITIME API: タイムアウトしました。")
        return None
    except requests.exceptions.RequestException as e:
        st.warning(f"⚠️ NAVITIME API 通信エラー: {e}")
        return None
    except Exception as e:
        st.warning(f"⚠️ NAVITIME API 処理エラー: {e}")
        return None


# ============================================================
# STEP 3: パターン生成と比較ロジック
# ============================================================

def build_best_route_patterns(current_st_name: str, ai_info: dict, navitime_route: dict,
                              headcount: int, work_days: int,
                              no_rental: bool = False, no_taxi: bool = False):
    """
    Args:
        no_rental: Trueの場合、レンタカーパターンを除外
        no_taxi: Trueの場合、タクシーパターンを除外
    """
    patterns = []

    if navitime_route:
        has_flight = navitime_route["has_flight"]
        time_min = navitime_route["time_min"]

        dest_lat, dest_lon = ai_info.get("dest_lat"), ai_info.get("dest_lon")
        e_lat, e_lon = navitime_route.get("end_station_lat"), navitime_route.get("end_station_lon")

        if e_lat and e_lon and dest_lat and dest_lon:
            base_taxi_km = haversine_km(e_lat, e_lon, dest_lat, dest_lon) * 1.3
            if base_taxi_km <= 1.2:
                base_taxi_km = 0.0
            else:
                base_taxi_km = max(base_taxi_km, 1.5)
        else:
            if navitime_route["last_walk_min"] <= TAXI_WALK_THRESHOLD_MIN:
                base_taxi_km = 0.0
            else:
                base_taxi_km = max(navitime_route["last_walk_min"] * 0.08, 1.5)

        if has_flight:
            selected_mode = "flight"
            flight_fare = navitime_route["flight_fare"]
            access_train_fare = navitime_route["access_train_fare"]

            origin_airport = navitime_route["flight_start"] or "出発空港"
            airport_name = navitime_route["flight_end"] or "到着空港"
            end_st = navitime_route["end_station"]

            if airport_name == end_st:
                access_str = f"{current_st_name} ➔ {origin_airport}"
            else:
                access_str = f"{current_st_name} ➔ {origin_airport} / {airport_name} ➔ {end_st}"

            display_route_str = f"{current_st_name} ➔ {origin_airport} ➔ {airport_name} ➔ {end_st} ➔ 目的地"

            route_dict = {
                "train": "-",
                "flight": f"{origin_airport} ➔ {airport_name}",
                "access": access_str,
                "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "現地徒歩",
                "rental": f"{end_st} ➔ 目的地周辺"
            }
            route_title_prefix = f"✈️ 飛行機ルート ({airport_name}経由 {end_st}着)"
            is_ai_fare = False

        else:
            selected_mode = "train"
            end_st = navitime_route["end_station"]
            route_dict = {
                "train": f"{current_st_name} ➔ {end_st}",
                "flight": "-",
                "access": "-",
                "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "現地徒歩",
                "rental": f"{end_st} ➔ 目的地周辺"
            }
            route_title_prefix = f"🚄 新幹線/電車ルート ({end_st}着)"
            display_route_str = f"{current_st_name} ➔ {end_st} ➔ 目的地"
            is_ai_fare = False

    else:
        # NAVITIMEをスキップ（離島）→ STEP 1.5 で取得した運賃を使用
        selected_mode = "flight"
        airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
        airport_lat = ai_info.get("airport_lat", 0)
        airport_lon = ai_info.get("airport_lon", 0)
        dest_lat = ai_info.get("dest_lat", 0)
        dest_lon = ai_info.get("dest_lon", 0)

        airport_to_dest_km = 15.0
        if airport_lat and airport_lon and dest_lat and dest_lon:
            airport_to_dest_km = max(
                haversine_km(airport_lat, airport_lon, dest_lat, dest_lon) * 1.3, 1.0
            )

        time_min = int(60 + ai_info.get("flight_time_min", 150) + (airport_to_dest_km / 40.0 * 60))

        # AI相場検索の運賃に1.3倍マージンを適用（flex料金ベース）
        flight_fare = int(ai_info.get("flight_fare_estimate", 60000) * 1.3)

        origin_airport = "伊丹空港" if "淀屋橋" in current_st_name else "羽田空港"
        access_train_fare = 1500 if "大宮" in current_st_name else 500

        route_dict = {
            "train": "-",
            "flight": f"{origin_airport} ➔ {airport_name}",
            "access": f"{current_st_name} ➔ {origin_airport}",
            "taxi": f"{airport_name} ↔ 目的地",
            "rental": f"{airport_name} ➔ 目的地周辺"
        }
        base_taxi_km = airport_to_dest_km
        route_title_prefix = f"✈️ 飛行機ルート ({airport_name}利用)"
        display_route_str = f"{current_st_name} ➔ {origin_airport} ➔ {airport_name} ➔ 目的地"
        is_ai_fare = True

    # 宿泊判定
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

    # 費用計算
    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)
    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    rental_car_total = round_up_1000(RENTAL_CAR_COST_PER_DAY * rental_days)

    # タクシー計算
    taxi_one_way = base_taxi_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1
    taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

    # ★ 交通費の基本部分（飛行機 or 電車）
    base_transport_taxi = {}
    base_transport_rental = {}

    if selected_mode == "flight":
        base_transport_taxi["航空券費用(往復・人数分・flex)"] = round_up_1000(flight_fare * 2 * headcount)
        base_transport_taxi["アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
        base_transport_rental["航空券費用(往復・人数分・flex)"] = round_up_1000(flight_fare * 2 * headcount)
        base_transport_rental["アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
    else:
        base_transport_taxi["電車・新幹線運賃(往復・人数分)"] = round_up_1000(navitime_route["total_fare"] * 2 * headcount)
        base_transport_rental["電車・新幹線運賃(往復・人数分)"] = round_up_1000(navitime_route["total_fare"] * 2 * headcount)

    # ★ タクシーパターン
    b_taxi = dict(base_transport_taxi)
    b_taxi[f"現地タクシー運賃(往復×{taxi_trips}回分)"] = taxi_total
    b_taxi["宿泊費"] = hotel_cost
    taxi_sum = sum(b_taxi.values())

    # ★ レンタカーパターン
    b_rental = dict(base_transport_rental)
    b_rental[f"レンタカー費用(12,000円×{rental_days}日)"] = rental_car_total
    b_rental["宿泊費"] = hotel_cost
    rental_sum = sum(b_rental.values())

    # ★ 除外オプションに基づいてパターンを選択
    if no_taxi and no_rental:
        # 両方除外 → 徒歩前提（タクシー・レンタカー費用なし）
        b_walk = dict(base_transport_taxi)
        b_walk["宿泊費"] = hotel_cost
        walk_sum = sum(b_walk.values())
        recommend_msg = "🏆 最安ルート (現地移動なし・徒歩前提)"
        final_breakdown = b_walk
        final_cost = walk_sum
        final_type = "walk"
        final_name = f"{route_title_prefix} (現地徒歩)"

    elif no_rental:
        # レンタカー除外 → タクシーのみ
        if taxi_total == 0:
            recommend_msg = "🏆 最安ルート (駅・空港から徒歩圏内)"
            final_name = f"{route_title_prefix} (現地徒歩)"
        else:
            recommend_msg = "🏆 最安ルート (タクシー利用)"
            final_name = f"{route_title_prefix} ＋ 現地タクシー"
        final_breakdown = b_taxi
        final_cost = taxi_sum
        final_type = "taxi"

    elif no_taxi:
        # タクシー除外 → レンタカーのみ
        recommend_msg = "🏆 最安ルート (レンタカー利用)"
        final_breakdown = b_rental
        final_cost = rental_sum
        final_type = "rental"
        final_name = f"{route_title_prefix} ＋ 現地レンタカー ({rental_days}日間)"

    else:
        # 両方有効 → 比較して安い方を選択
        if rental_sum < taxi_sum:
            diff = taxi_sum - rental_sum
            recommend_msg = f"🏆 最安ルート (レンタカー利用・タクシーより {diff:,} 円お得)"
            final_breakdown = b_rental
            final_cost = rental_sum
            final_type = "rental"
            final_name = f"{route_title_prefix} ＋ 現地レンタカー ({rental_days}日間)"
        else:
            if taxi_total == 0:
                recommend_msg = "🏆 最安ルート (駅・空港から徒歩圏内)"
                final_name = f"{route_title_prefix} (現地徒歩)"
            else:
                recommend_msg = "🏆 最安ルート (タクシー利用推奨)"
                final_name = f"{route_title_prefix} ＋ 現地タクシー"
            final_breakdown = b_taxi
            final_cost = taxi_sum
            final_type = "taxi"

    if is_ai_fare:
        recommend_msg += " / ⚠️ AI相場検索(運賃1.3倍マージン)"

    # ★ デバッグ: 費用計算の内訳を表示
    if DEBUG_MODE:
        with st.expander("🔍 [DEBUG] 費用計算の内訳", expanded=False):
            st.markdown(f"**selected_mode:** `{selected_mode}`")
            st.markdown(f"**time_min:** {time_min} 分 ({travel_hours:.1f} 時間)")
            st.markdown(f"**nights:** {nights} 泊")
            st.markdown(f"**base_taxi_km:** {base_taxi_km:.1f} km")
            st.markdown(f"**taxi_trips:** {taxi_trips} 回")
            st.markdown(f"**除外オプション:** レンタカー={'除外' if no_rental else '有効'}, タクシー={'除外' if no_taxi else '有効'}")
            st.markdown("---")
            if selected_mode == "flight":
                st.markdown(f"**flight_fare (片道・flex):** {flight_fare:,} 円")
                st.markdown(f"**access_train_fare (片道):** {access_train_fare:,} 円")
            else:
                st.markdown(f"**total_fare (片道):** {navitime_route['total_fare']:,} 円")
            st.markdown("---")
            if not no_taxi:
                st.markdown("**タクシーパターン:**")
                for k, v in b_taxi.items():
                    st.markdown(f"- {k}: **{v:,}** 円")
                st.markdown(f"- **合計: {taxi_sum:,} 円**")
                st.markdown("---")
            if not no_rental:
                st.markdown("**レンタカーパターン:**")
                for k, v in b_rental.items():
                    st.markdown(f"- {k}: **{v:,}** 円")
                st.markdown(f"- **合計: {rental_sum:,} 円**")
                st.markdown("---")
            st.markdown(f"**採用:** {final_type}")

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
# Streamlit UI（公開版）
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗", layout="wide")
st.title("🚗 交通費・出張見積もりアプリ")

# デバッグモード表示
if DEBUG_MODE:
    st.caption("🐛 デバッグモード ON — NAVITIMEレスポンスと計算内訳が表示されます")

# サーバー側でAPIキーを取得
gemini_api_key, navitime_api_key = get_api_keys()

col1, col2 = st.columns([2, 1])
with col1:
    address_input = st.text_input("目的地（住所や施設名）", "")
with col2:
    station_choice = st.selectbox("出発拠点", ["淀屋橋駅", "大宮駅", "両方比較"], index=0)

col_a, col_b = st.columns(2)
with col_a:
    headcount = st.number_input("作業人数（人）", min_value=1, value=1)
with col_b:
    work_days = st.number_input("現地作業日数（日）", min_value=1, value=2)

# ★ 交通手段除外チェックボックス
st.markdown("**🚫 使用しない交通手段:**")
col_c1, col_c2, col_c3, col_c4 = st.columns(4)
with col_c1:
    no_flight = st.checkbox("飛行機を使用しない", value=False)
with col_c2:
    no_shinkansen = st.checkbox("新幹線を使用しない", value=False)
with col_c3:
    no_rental = st.checkbox("レンタカーを使用しない", value=False)
with col_c4:
    no_taxi = st.checkbox("タクシーを使用しない", value=False)

st.markdown("---")

if st.button("🚀 最速出張見積もりを計算する", type="primary"):

    if not address_input.strip():
        st.warning("⚠️ 目的地を入力してください。")
        st.stop()

    # ★ 飛行機と新幹線の両方を除外した場合の警告
    if no_flight and no_shinkansen:
        st.info("💡 飛行機・新幹線の両方を除外しています。在来線のみでルートを検索します。")

    stations = (STATION_COORDS if station_choice == "両方比較"
                else {station_choice: STATION_COORDS[station_choice]})

    for current_st_name, (current_lat, current_lon) in stations.items():
        st.markdown(f"### 🚉 出発地: 【{current_st_name}】")

        # ── STEP 1: 目的地の基本分析（Google Search なし・軽量） ──
        with st.spinner(f"🤖 AIが {address_input} を分析中..."):
            ai_info = analyze_destination_with_gemini(address_input, gemini_api_key, current_st_name)

        if not ai_info.get("dest_lat") or not ai_info.get("dest_lon"):
            st.error("❌ 目的地の座標が取得できませんでした。住所を詳しく入力してください。")
            continue

        st.success(f"**📍 検索地点:** {ai_info.get('normalized_address', address_input)}")

        # デバッグ: Gemini分析結果
        if DEBUG_MODE:
            with st.expander("🔍 [DEBUG] Gemini分析結果 (STEP 1)", expanded=False):
                st.json(ai_info)

        train_route = None
        if ai_info.get("is_island_or_remote", False) and not no_flight:
            st.info("💡 **海を渡る離島・遠隔地と判定されました。NAVITIME検索をスキップし、飛行機ルートを適用します。**")

            # ── STEP 1.5: 離島の場合のみ Google Search で運賃検索 ──
            airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
            with st.spinner(f"✈️ {airport_name} への運賃をGoogle検索中..."):
                fare_info = search_flight_fare_with_gemini(
                    address_input, airport_name, gemini_api_key, current_st_name
                )
            ai_info.update(fare_info)

            # デバッグ: 運賃検索結果
            if DEBUG_MODE:
                with st.expander("🔍 [DEBUG] 運賃検索結果 (STEP 1.5)", expanded=False):
                    st.json(fare_info)

            if not ai_info.get("airport_lat"):
                ai_info["airport_lat"] = ai_info["dest_lat"] if ai_info["dest_lat"] else 24.3964
                ai_info["airport_lon"] = ai_info["dest_lon"] if ai_info["dest_lon"] else 124.2450

        elif ai_info.get("is_island_or_remote", False) and no_flight:
            # 離島だが飛行機除外 → 警告を出してNAVITIMEで検索（船便等）
            st.warning("⚠️ 離島ですが飛行機を除外しています。船便等のルートを検索しますが、見つからない場合があります。")
            with st.spinner("🧭 NAVITIMEでルートを検索中..."):
                train_route = get_navitime_fastest_route(
                    current_lat, current_lon, ai_info["dest_lat"], ai_info["dest_lon"],
                    navitime_api_key, no_flight=no_flight, no_shinkansen=no_shinkansen
                )

        else:
            # ── STEP 2: 本土の場合は NAVITIME で検索 ──
            with st.spinner("🧭 NAVITIMEで全国対応の最速ルートを検索中..."):
                train_route = get_navitime_fastest_route(
                    current_lat, current_lon, ai_info["dest_lat"], ai_info["dest_lon"],
                    navitime_api_key, no_flight=no_flight, no_shinkansen=no_shinkansen
                )

        # ── STEP 3: パターン生成・比較 ──
        patterns = build_best_route_patterns(
            current_st_name, ai_info, train_route, headcount, work_days,
            no_rental=no_rental, no_taxi=no_taxi
        )

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

                excel_bytes = create_excel_report(
                    p, ai_info.get('normalized_address', address_input), headcount, work_days
                )
                st.download_button(
                    label="📥 Excel見積書をダウンロード",
                    data=excel_bytes,
                    file_name=f"交通費見積書_{current_st_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{current_st_name}_{p['type']}"
                )
        st.markdown("---")

