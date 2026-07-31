
# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（Streamlit Cloud公開版）

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

MAX_RETRIES = 3
RETRY_WAIT_SECONDS = [5, 15, 30]
RETRYABLE_STATUS_CODES = {429, 503}

DEBUG_MODE = True


# ============================================================
# APIキー取得
# ============================================================

def get_api_keys():
    try:
        gemini_key = st.secrets["api_keys"]["gemini"]
        navitime_key = st.secrets["api_keys"]["navitime"]
        return gemini_key, navitime_key
    except KeyError:
        st.error(
            "APIキーが設定されていません。\n\n"
            "Streamlit Cloud の Settings → Secrets に以下の形式で登録してください:\n\n"
            "```toml\n[api_keys]\ngemini = \"YOUR_GEMINI_API_KEY\"\nnavitime = \"YOUR_NAVITIME_RAPIDAPI_KEY\"\n```"
        )
        st.stop()
        return "", ""


# ============================================================
# Gemini API リトライラッパー
# ============================================================

def call_gemini_with_retry(client, model, contents, config, retry_status_placeholder=None):
    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=model, contents=contents, config=config,
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
                    f"サーバー混雑中... {wait_sec}秒後にリトライします（{attempt + 1}/{MAX_RETRIES}回目）"
                )
            time.sleep(wait_sec)
            if retry_status_placeholder:
                retry_status_placeholder.empty()
    raise last_error


# ============================================================
# ユーティリティ関数
# ============================================================

def round_up_1000(amount):
    if amount <= 0:
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


def guess_airport_from_address(address):
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


def parse_json_from_text(text):
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
# Excelファイル出力関数（社内フォーマット完全再現）
# ============================================================

def create_excel_report(pattern_data, address, headcount, work_days, origin_station):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交通費"

    # --- スタイル定義 ---
    hdr_fill = PatternFill(start_color="D9531E", end_color="D9531E", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    font_normal = Font(name="メイリオ", size=11)
    font_bold = Font(name="メイリオ", size=11, bold=True)
    font_hdr = Font(name="メイリオ", size=11, bold=True, color="FFFFFF")

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    border_hdr = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_fmt = '#,##0_);[Red](#,##0)'

    # --- 列幅 ---
    col_widths = {
        'A': 2.94, 'B': 8.09, 'C': 4.17, 'D': 13.0, 'E': 13.0,
        'F': 13.0, 'G': 5.64, 'H': 11.4, 'I': 6.86, 'J': 9.68,
        'K': 5.15, 'L': 4.17, 'M': 8.95, 'N': 8.21, 'O': 10.17,
        'P': 8.09, 'Q': 10.54, 'R': 17.16, 'S': 27.95
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    for r in [2, 3, 6, 8, 9, 10, 11, 12, 13, 14]:
        ws.row_dimensions[r].height = 18.75

    ws.merge_cells('C8:G8')
    ws.merge_cells('H8:L8')

    # --- データ取得 ---
    routes = pattern_data.get("routes", {})
    breakdown = pattern_data.get("breakdown", {})
    time_hours = round(pattern_data.get("time_min", 0) / 60.0, 1)
    stay_note = pattern_data.get("note", "")

    def parse_route(route_str):
        if not route_str or route_str == "-":
            return "", ""
        parts = re.split(r'\s*[➔→/]\s*', route_str)
        if len(parts) >= 2:
            return parts[0].strip(), parts[-1].strip()
        return route_str.strip(), ""

    train_route = routes.get("train", "-")
    flight_route = routes.get("flight", "-")
    access_route = routes.get("access", "-")

    train_from, train_to = parse_route(train_route)
    flight_from, flight_to = parse_route(flight_route)
    access_from, access_to = parse_route(access_route)

    def get_breakdown_amount(key_contains):
        for k, v in breakdown.items():
            if key_contains in k:
                return int(v)
        return 0

    train_total = get_breakdown_amount("電車・新幹線運賃")
    flight_total = get_breakdown_amount("航空券費用")
    access_total = get_breakdown_amount("アクセス電車運賃")
    rental_total = get_breakdown_amount("レンタカー費用")
    taxi_total = get_breakdown_amount("タクシー運賃")
    hotel_total = get_breakdown_amount("宿泊費")

    is_train_used = train_total > 0
    is_flight_used = flight_total > 0
    is_access_used = access_total > 0
    is_rental_used = rental_total > 0
    is_taxi_used = taxi_total > 0
    is_hotel_used = hotel_total > 0

    # N列の値（往復1人分の実費用）
    train_n = int(train_total / headcount) if headcount > 0 and is_train_used else 0
    flight_n = int(flight_total / headcount) if headcount > 0 and is_flight_used else 0
    access_n = int(access_total / headcount) if headcount > 0 and is_access_used else 0
    rental_n = RENTAL_CAR_COST_PER_DAY if is_rental_used else 0
    taxi_n = int(taxi_total) if is_taxi_used else 0
    hotel_n = HOTEL_COST_PER_NIGHT_PER_PERSON if is_hotel_used else 0

    travel_days = 1 if time_hours >= 4 else 0

    nights = 0
    match_nights = re.search(r'(\d+)泊', stay_note)
    if match_nights:
        nights = int(match_nights.group(1))
    elif is_hotel_used:
        nights = work_days

    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    flight_origin_airport = "伊丹空港" if "淀屋橋" in origin_station else "羽田空港"

    # ===== 設置設定作業セクション (Row 1〜15) =====

    ws['B1'] = "設置設定作業"
    ws['B1'].font = font_bold
    ws['H1'] = "基本大阪と埼玉のどちらか近い方から算出するが、TIS、CSI案件のみ全て大阪から算出する。"
    ws['H1'].font = font_normal
    ws['S1'] = "作業日数"
    ws['S1'].font = font_normal

    ws['H2'] = "大阪：淀屋橋駅、埼玉：大宮駅"
    ws['H2'].font = font_normal
    ws['S2'] = "2日（作業日1日、予備日：1日）"
    ws['S2'].font = font_normal

    ws['B3'] = "住所"
    ws['B3'].font = font_normal
    ws['C3'] = address
    ws['C3'].font = font_normal
    ws['C3'].fill = yellow_fill
    ws['S3'] = "4日（作業日3日、予備日：1日）"
    ws['S3'].font = font_normal

    ws['B4'] = "人数"
    ws['B4'].font = font_normal
    ws['C4'] = headcount
    ws['C4'].font = font_normal
    ws['D4'] = "人"
    ws['D4'].font = font_normal
    ws['S4'] = "5日（作業日4日、予備日：1日）"
    ws['S4'].font = font_normal

    ws['B5'] = "作業"
    ws['B5'].font = font_normal
    ws['C5'] = work_days
    ws['C5'].font = font_normal
    ws['D5'] = "日"
    ws['D5'].font = font_normal

    ws['B6'] = "移動"
    ws['B6'].font = font_normal
    ws['C6'] = "=IF(R15>=4,1,0)"
    ws['C6'].font = font_normal
    ws['D6'] = "日"
    ws['D6'].font = font_normal
    ws['E6'] = "※移動に2.5h以上かかる場合は宿泊想定／移動に4h以上かかる場合又は飛行機を利用の場合はさらに前泊か後泊想定"
    ws['E6'].font = font_normal

    ws['M7'] = "※最大費用の経路を想定"
    ws['M7'].font = font_normal

    # Row 8: ヘッダー
    hdr_cells = {
        'B': 'チェック', 'C': '項目', 'H': '経路',
        'M': '調整費用', 'N': '実費用', 'O': '日数・泊数',
        'P': '人数', 'Q': '合計', 'R': '片路移動時間（h）'
    }
    for col_letter, label in hdr_cells.items():
        cell = ws[f'{col_letter}8']
        cell.value = label
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr

    for col in range(4, 8):
        cell = ws.cell(row=8, column=col)
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr
    for col in range(9, 13):
        cell = ws.cell(row=8, column=col)
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr

    # --- データ行 Row 9〜14 ---
    rows_data = [
        {
            "check": "✓" if is_train_used else "-",
            "item": "電車・新幹線（往復）　",
            "h": origin_station if is_train_used else "",
            "i": "駅->",
            "j": train_to if is_train_used else "",
            "k": "駅",
            "n": train_n, "o": 1, "p": "=C4",
            "r": time_hours if is_train_used else "",
            "s": ""
        },
        {
            "check": "✓" if is_flight_used else "-",
            "item": "飛行機（往復）　",
            "h": flight_from if is_flight_used else flight_origin_airport,
            "i": "空港->",
            "j": flight_to if is_flight_used else "",
            "k": "空港",
            "n": flight_n, "o": 1, "p": "=C4",
            "r": time_hours if is_flight_used and not is_train_used else "",
            "s": "ANA/フレックスで試算"
        },
        {
            "check": "✓" if is_access_used else "-",
            "item": "電車・新幹線（往復）　",
            "h": access_from if is_access_used else "",
            "i": "駅->",
            "j": access_to if is_access_used else "",
            "k": "駅",
            "n": access_n, "o": 1, "p": "=C4",
            "r": "",
            "s": ""
        },
        {
            "check": "✓" if is_rental_used else "-",
            "item": "レンタカー",
            "h": "", "i": "", "j": "", "k": "",
            "n": rental_n, "o": rental_days if is_rental_used else 1, "p": "1",
            "r": 1 if is_rental_used else "",
            "s": "使用基準決める"
        },
        {
            "check": "✓" if is_taxi_used else "-",
            "item": "タクシー（往復）",
            "h": "", "i": "", "j": "", "k": "",
            "n": taxi_n, "o": "=C5+C6", "p": "1",
            "r": "",
            "s": "バス/使用基準決める"
        },
        {
            "check": "✓" if is_hotel_used else "-",
            "item": "宿泊",
            "h": "", "i": "", "j": "", "k": "",
            "n": hotel_n, "o": nights if is_hotel_used else 1, "p": "=C4",
            "r": "-",
            "s": ""
        },
    ]

    for idx, rd in enumerate(rows_data):
        row_num = 9 + idx

        ws.cell(row=row_num, column=2, value=rd["check"]).font = font_normal
        ws.cell(row=row_num, column=2).alignment = align_center
        ws.cell(row=row_num, column=2).border = border_data
        if rd["check"] == "✓":
            ws.cell(row=row_num, column=2).fill = yellow_fill

        ws.cell(row=row_num, column=3, value=rd["item"]).font = font_normal
        ws.cell(row=row_num, column=3).alignment = align_left
        ws.cell(row=row_num, column=3).border = border_data

        ws.cell(row=row_num, column=8, value=rd["h"]).font = font_normal
        ws.cell(row=row_num, column=8).border = border_data
        if rd["h"] and rd["check"] == "✓":
            ws.cell(row=row_num, column=8).fill = yellow_fill

        ws.cell(row=row_num, column=9, value=rd["i"]).font = font_normal
        ws.cell(row=row_num, column=9).alignment = align_center
        ws.cell(row=row_num, column=9).border = border_data

        ws.cell(row=row_num, column=10, value=rd["j"]).font = font_normal
        ws.cell(row=row_num, column=10).border = border_data
        if rd["j"] and rd["check"] == "✓":
            ws.cell(row=row_num, column=10).fill = yellow_fill

        ws.cell(row=row_num, column=11, value=rd["k"]).font = font_normal
        ws.cell(row=row_num, column=11).alignment = align_center
        ws.cell(row=row_num, column=11).border = border_data

        # M列: 調整費用
        if rd["item"] in ("レンタカー", "宿泊"):
            ws.cell(row=row_num, column=13, value=rd["n"]).font = font_normal
        else:
            ws.cell(row=row_num, column=13, value=f"=ROUNDUP(N{row_num},-3)").font = font_normal
        ws.cell(row=row_num, column=13).alignment = align_center
        ws.cell(row=row_num, column=13).border = border_data
        ws.cell(row=row_num, column=13).number_format = num_fmt

        # N列: 実費用
        ws.cell(row=row_num, column=14, value=rd["n"]).font = font_normal
        ws.cell(row=row_num, column=14).alignment = align_center
        ws.cell(row=row_num, column=14).border = border_data
        ws.cell(row=row_num, column=14).number_format = num_fmt
        if rd["check"] == "✓":
            ws.cell(row=row_num, column=14).fill = yellow_fill

        # O列: 日数
        ws.cell(row=row_num, column=15, value=rd["o"]).font = font_normal
        ws.cell(row=row_num, column=15).alignment = align_center
        ws.cell(row=row_num, column=15).border = border_data
        if rd["check"] == "✓" and isinstance(rd["o"], int):
            ws.cell(row=row_num, column=15).fill = yellow_fill

        # P列: 人数
        ws.cell(row=row_num, column=16, value=rd["p"]).font = font_normal
        ws.cell(row=row_num, column=16).alignment = align_center
        ws.cell(row=row_num, column=16).border = border_data

        # Q列: 合計
        q_formula = f'=IF(B{row_num}="✓",SUM(M{row_num}*O{row_num}*P{row_num}),IF(B{row_num}="-",0,"確認"))'
        ws.cell(row=row_num, column=17, value=q_formula).font = font_normal
        ws.cell(row=row_num, column=17).alignment = align_center
        ws.cell(row=row_num, column=17).border = border_data
        ws.cell(row=row_num, column=17).number_format = num_fmt

        # R列: 移動時間
        ws.cell(row=row_num, column=18, value=rd["r"]).font = font_normal
        ws.cell(row=row_num, column=18).alignment = align_center
        ws.cell(row=row_num, column=18).border = border_data
        ws.cell(row=row_num, column=18).number_format = num_fmt
        if rd["r"] and rd["r"] != "-" and rd["check"] == "✓":
            ws.cell(row=row_num, column=18).fill = yellow_fill

        # S列: 備考
        if rd["s"]:
            ws.cell(row=row_num, column=19, value=rd["s"]).font = font_normal

    # Row 15: 合計
    ws.cell(row=15, column=17, value="=SUM(Q9:Q14)").font = font_bold
    ws.cell(row=15, column=17).number_format = num_fmt
    ws.cell(row=15, column=18, value='=SUMIF(B9:B14,"✓",R9:R14)').font = font_bold
    ws.cell(row=15, column=18).number_format = num_fmt

    # ===== 下見作業セクション (Row 17〜31) =====

    ws['B17'] = "下見作業"
    ws['B17'].font = font_bold

    ws['B19'] = "住所"
    ws['B19'].font = font_normal
    ws['C19'] = "=C3"
    ws['C19'].font = font_normal

    ws['B20'] = "人数"
    ws['B20'].font = font_normal
    ws['C20'] = 1
    ws['C20'].font = font_normal
    ws['D20'] = "人"
    ws['D20'].font = font_normal

    ws['B21'] = "作業"
    ws['B21'].font = font_normal
    ws['C21'] = 1
    ws['C21'].font = font_normal
    ws['D21'] = "日"
    ws['D21'].font = font_normal

    ws['B22'] = "移動"
    ws['B22'].font = font_normal
    ws['C22'] = "=IF(R31>4,1,0)"
    ws['C22'].font = font_normal
    ws['D22'] = "日"
    ws['D22'].font = font_normal
    ws['E22'] = "※移動に4h以上かかる場合又は飛行機を利用の場合はさらに前泊か後泊想定"
    ws['E22'].font = font_normal

    ws['M23'] = "※最大費用の経路を想定"
    ws['M23'].font = font_normal

    # Row 24: ヘッダー
    ws.merge_cells('C24:G24')
    ws.merge_cells('H24:L24')
    for col_letter, label in hdr_cells.items():
        cell = ws[f'{col_letter}24']
        cell.value = label
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr
    for col in range(4, 8):
        cell = ws.cell(row=24, column=col)
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr
    for col in range(9, 13):
        cell = ws.cell(row=24, column=col)
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr

    # Row 25〜30: 下見データ
    shitami_rows = [
        {"check": "-", "item": "電車・新幹線（往復）　", "h": "=H9", "i": "駅->", "j": "=J9", "k": "駅",
         "m": "=ROUNDUP(N25,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R9"},
        {"check": "-", "item": "飛行機（往復）　", "h": flight_origin_airport, "i": "空港->", "j": "", "k": "空港",
         "m": "=ROUNDUP(N26,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R10"},
        {"check": "-", "item": "電車・新幹線（往復）　", "h": "", "i": "駅->", "j": "", "k": "駅",
         "m": "=ROUNDUP(N27,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R11"},
        {"check": "-", "item": "レンタカー", "h": "", "i": "", "j": "", "k": "",
         "m": 12000, "n": "", "o": "=C21+C22", "p": "1", "r": "=R12"},
        {"check": "-", "item": "タクシー（往復）", "h": "", "i": "", "j": "", "k": "",
         "m": "=ROUNDUP(N29,-3)", "n": "", "o": "=C21+C22", "p": "1", "r": "=R13"},
        {"check": "-", "item": "宿泊", "h": "", "i": "", "j": "", "k": "",
         "m": 20000, "n": 20000, "o": "=C21+C22-1", "p": "=C20", "r": "-"},
    ]

    for idx, rd in enumerate(shitami_rows):
        row_num = 25 + idx

        ws.cell(row=row_num, column=2, value=rd["check"]).font = font_normal
        ws.cell(row=row_num, column=2).alignment = align_center
        ws.cell(row=row_num, column=2).border = border_data

        ws.cell(row=row_num, column=3, value=rd["item"]).font = font_normal
        ws.cell(row=row_num, column=3).alignment = align_left
        ws.cell(row=row_num, column=3).border = border_data

        ws.cell(row=row_num, column=8, value=rd["h"]).font = font_normal
        ws.cell(row=row_num, column=8).border = border_data

        ws.cell(row=row_num, column=9, value=rd["i"]).font = font_normal
        ws.cell(row=row_num, column=9).alignment = align_center
        ws.cell(row=row_num, column=9).border = border_data

        ws.cell(row=row_num, column=10, value=rd["j"]).font = font_normal
        ws.cell(row=row_num, column=10).border = border_data

        ws.cell(row=row_num, column=11, value=rd["k"]).font = font_normal
        ws.cell(row=row_num, column=11).alignment = align_center
        ws.cell(row=row_num, column=11).border = border_data

        ws.cell(row=row_num, column=13, value=rd["m"]).font = font_normal
        ws.cell(row=row_num, column=13).alignment = align_center
        ws.cell(row=row_num, column=13).border = border_data
        ws.cell(row=row_num, column=13).number_format = num_fmt

        if rd["n"] != "":
            ws.cell(row=row_num, column=14, value=rd["n"]).font = font_normal
            ws.cell(row=row_num, column=14).border = border_data
            ws.cell(row=row_num, column=14).number_format = num_fmt

        ws.cell(row=row_num, column=15, value=rd["o"]).font = font_normal
        ws.cell(row=row_num, column=15).alignment = align_center
        ws.cell(row=row_num, column=15).border = border_data

        ws.cell(row=row_num, column=16, value=rd["p"]).font = font_normal
        ws.cell(row=row_num, column=16).alignment = align_center
        ws.cell(row=row_num, column=16).border = border_data

        q_formula = f'=IF(B{row_num}="✓",SUM(M{row_num}*O{row_num}*P{row_num}),IF(B{row_num}="-",0,"確認"))'
        ws.cell(row=row_num, column=17, value=q_formula).font = font_normal
        ws.cell(row=row_num, column=17).alignment = align_center
        ws.cell(row=row_num, column=17).border = border_data
        ws.cell(row=row_num, column=17).number_format = num_fmt

        ws.cell(row=row_num, column=18, value=rd["r"]).font = font_normal
        ws.cell(row=row_num, column=18).alignment = align_center
        ws.cell(row=row_num, column=18).border = border_data
        ws.cell(row=row_num, column=18).number_format = num_fmt

    # Row 31: 下見合計
    ws.cell(row=31, column=17, value="=SUM(Q25:Q30)").font = font_bold
    ws.cell(row=31, column=17).number_format = num_fmt
    ws.cell(row=31, column=18, value='=SUMIF(B25:B30,"✓",R25:R30)').font = font_bold
    ws.cell(row=31, column=18).number_format = num_fmt

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による目的地分析
# ============================================================

def geocode_fallback(address):
    try:
        res = requests.get(GSI_GEOCODE_URL, params={"q": address}, timeout=5)
        if res.status_code == 200 and res.json():
            lon, lat = res.json()[0]["geometry"]["coordinates"]
            return lat, lon
    except Exception:
        pass
    return None, None


def analyze_destination_with_gemini(raw_address, gemini_key, origin_name):
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
  "is_island_or_remote": 対象が海を渡る完全な離島ならtrue。北海道・本州・四国・九州の本土はfalse,
  "nearest_airport_name": "離島の場合のみ最寄り空港名(本土の場合は空文字)"
}}
"""
    islands = ["沖縄", "奄美", "沖永良部", "石垣", "宮古", "屋久島", "種子島", "徳之島", "与論", "久米島", "喜界"]
    is_island_guess = any(island in raw_address for island in islands)

    fallback_data = {
        "normalized_address": raw_address,
        "is_island_or_remote": is_island_guess,
        "nearest_airport_name": guess_airport_from_address(raw_address) if is_island_guess else "",
        "dest_lat": None, "dest_lon": None,
    }

    retry_placeholder = st.empty()
    try:
        config = types.GenerateContentConfig(temperature=0.1)
        text = call_gemini_with_retry(client=client, model=GEMINI_MODEL, contents=prompt,
                                      config=config, retry_status_placeholder=retry_placeholder)
        if text:
            parsed_data = parse_json_from_text(text)
            if parsed_data:
                fallback_data.update(parsed_data)
    except Exception as e:
        st.warning(f"Gemini API呼び出し失敗（フォールバック値を使用）: {e}")
    finally:
        retry_placeholder.empty()

    if not fallback_data.get("dest_lat") or not fallback_data.get("dest_lon"):
        lat, lon = geocode_fallback(fallback_data["normalized_address"])
        fallback_data["dest_lat"] = lat
        fallback_data["dest_lon"] = lon

    if fallback_data.get("is_island_or_remote"):
        guessed_airport = guess_airport_from_address(fallback_data["normalized_address"])
        if not fallback_data.get("nearest_airport_name") or fallback_data["nearest_airport_name"] == "最寄り空港":
            fallback_data["nearest_airport_name"] = guessed_airport

    return fallback_data


# ============================================================
# STEP 1.5: 離島の場合のみ Gemini + Google Search
# ============================================================

def search_flight_fare_with_gemini(raw_address, airport_name, gemini_key, origin_name):
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
        "airport_lat": None, "airport_lon": None,
        "flight_fare_estimate": 60000, "flight_time_min": 180,
    }
    retry_placeholder = st.empty()
    try:
        google_search_tool = types.Tool(google_search=types.GoogleSearch())
        config = types.GenerateContentConfig(tools=[google_search_tool], temperature=0.1)
        text = call_gemini_with_retry(client=client, model=GEMINI_MODEL, contents=prompt,
                                      config=config, retry_status_placeholder=retry_placeholder)
        if text:
            parsed_data = parse_json_from_text(text)
            if parsed_data:
                fallback_data.update(parsed_data)
    except Exception as e:
        st.warning(f"運賃検索失敗（フォールバック値を使用）: {e}")
    finally:
        retry_placeholder.empty()
    return fallback_data


# ============================================================
# STEP 2: NAVITIME API
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, navitime_key,
                               no_flight=False, no_shinkansen=False):
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
            st.warning(f"NAVITIME API エラー: HTTP {res.status_code}")
            return None

        data = res.json()
        items = data.get("items", [])
        if not items:
            st.warning("NAVITIME: ルートが見つかりませんでした。")
            return None

        fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
        time_min = fastest_item.get("summary", {}).get("move", {}).get("time", 0)
        move_info = fastest_item.get("summary", {}).get("move", {})

        move_types = move_info.get("move_type", [])
        has_superexpress = "superexpress_train" in move_types

        fare_dict = move_info.get("fare", {})
        if isinstance(fare_dict, dict):
            fare_unit_0 = fare_dict.get("unit_0", 0)
            fare_unit_3 = fare_dict.get("unit_3", 0)
        else:
            fare_unit_0 = 0
            fare_unit_3 = 0

        if has_superexpress:
            total_fare = int(fare_unit_0 + fare_unit_3)
        else:
            total_fare = int(fare_unit_0)

        has_flight = False
        flight_fare = 0
        start_station_name = None
        end_station_name = None
        end_station_lat = None
        end_station_lon = None
        flight_start_name = None
        flight_end_name = None

        sections = fastest_item.get("sections", [])
        debug_sections = []

        for i, sec in enumerate(sections):
            m_type = sec.get("move", "")
            sec_type = sec.get("type", "")

            if DEBUG_MODE:
                debug_sections.append({
                    "index": i, "type": sec_type, "move": m_type,
                    "name": sec.get("name", ""), "time": sec.get("time", ""),
                    "line_name": sec.get("line_name", ""),
                })

            if sec_type == "point":
                if not start_station_name:
                    start_station_name = sec.get("name")
                p_name = sec.get("name", "").lower()
                if "goal" not in p_name:
                    end_station_name = sec.get("name")
                    if "coord" in sec:
                        end_station_lat = sec["coord"].get("lat")
                        end_station_lon = sec["coord"].get("lon")

            if sec_type == "move" and m_type.lower() in FLIGHT_MOVE_TYPES:
                has_flight = True
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

        if DEBUG_MODE:
            with st.expander("🔍 [DEBUG] NAVITIMEレスポンス詳細", expanded=False):
                st.json(params)
                st.json(move_info)
                st.markdown(f"**move_type:** `{move_types}` | **has_superexpress:** `{has_superexpress}`")
                st.markdown(f"**採用運賃:** {total_fare:,} 円")
                if has_flight:
                    st.markdown(f"**航空券(flex):** {flight_fare:,} | **電車:** {access_train_fare:,}")
                for ds in debug_sections:
                    st.text(f"[{ds['index']}] {ds['type']}/{ds['move']} name={ds['name']} time={ds['time']}")
                st.json(result)

        return result

    except Exception as e:
        st.warning(f"NAVITIME API エラー: {e}")
        return None


# ============================================================
# STEP 3: パターン生成と比較ロジック
# ============================================================

def build_best_route_patterns(current_st_name, ai_info, navitime_route,
                              headcount, work_days,
                              no_rental=False, no_taxi=False):
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
            route_title_prefix = f"✈️ 飛行機ルート ({airport_name}経由)"
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
            route_title_prefix = f"🚄 電車ルート ({end_st}着)"
            display_route_str = f"{current_st_name} ➔ {end_st} ➔ 目的地"
            is_ai_fare = False
    else:
        selected_mode = "flight"
        airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
        airport_lat = ai_info.get("airport_lat", 0)
        airport_lon = ai_info.get("airport_lon", 0)
        dest_lat = ai_info.get("dest_lat", 0)
        dest_lon = ai_info.get("dest_lon", 0)

        airport_to_dest_km = 15.0
        if airport_lat and airport_lon and dest_lat and dest_lon:
            airport_to_dest_km = max(haversine_km(airport_lat, airport_lon, dest_lat, dest_lon) * 1.3, 1.0)

        time_min = int(60 + ai_info.get("flight_time_min", 150) + (airport_to_dest_km / 40.0 * 60))
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
        stay_note = "移動時間や飛行機利用のため、前泊/後泊想定（" + str(nights) + "泊）"
    elif travel_hours >= 2.5:
        nights = work_days
        stay_note = "移動2.5時間以上の為宿泊想定（" + str(nights) + "泊）"
    else:
        nights = max(work_days - 1, 0)
        stay_note = "日帰り/標準宿泊（" + str(nights) + "泊）"

    # 費用計算
    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)
    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    rental_car_total = round_up_1000(RENTAL_CAR_COST_PER_DAY * rental_days)

    taxi_one_way = base_taxi_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1
    taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

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

    b_taxi = dict(base_transport_taxi)
    if taxi_total > 0:
        b_taxi["現地タクシー運賃(往復×" + str(taxi_trips) + "回分)"] = taxi_total
    b_taxi["宿泊費"] = hotel_cost
    taxi_sum = sum(b_taxi.values())

    b_rental = dict(base_transport_rental)
    b_rental["レンタカー費用(12,000円×" + str(rental_days) + "日)"] = rental_car_total
    b_rental["宿泊費"] = hotel_cost
    rental_sum = sum(b_rental.values())

    # 除外オプション
    if no_taxi and no_rental:
        b_walk = dict(base_transport_taxi)
        b_walk["宿泊費"] = hotel_cost
        walk_sum = sum(b_walk.values())
        recommend_msg = "最安ルート (現地移動なし・徒歩前提)"
        final_breakdown = b_walk
        final_cost = walk_sum
        final_type = "walk"
        final_name = route_title_prefix + " (現地徒歩)"
    elif no_rental:
        if taxi_total == 0:
            recommend_msg = "最安ルート (駅・空港から徒歩圏内)"
            final_name = route_title_prefix + " (現地徒歩)"
        else:
            recommend_msg = "最安ルート (タクシー利用)"
            final_name = route_title_prefix + " ＋ 現地タクシー"
        final_breakdown = b_taxi
        final_cost = taxi_sum
        final_type = "taxi"
    elif no_taxi:
        recommend_msg = "最安ルート (レンタカー利用)"
        final_breakdown = b_rental
        final_cost = rental_sum
        final_type = "rental"
        final_name = route_title_prefix + " ＋ レンタカー"
    else:
        if rental_sum < taxi_sum:
            diff = taxi_sum - rental_sum
            recommend_msg = "最安ルート (レンタカー利用・タクシーより " + f"{diff:,}" + " 円お得)"
            final_breakdown = b_rental
            final_cost = rental_sum
            final_type = "rental"
            final_name = route_title_prefix + " ＋ レンタカー"
        else:
            if taxi_total == 0:
                recommend_msg = "最安ルート (駅・空港から徒歩圏内)"
                final_name = route_title_prefix + " (現地徒歩)"
            else:
                recommend_msg = "最安ルート (タクシー利用推奨)"
                final_name = route_title_prefix + " ＋ 現地タクシー"
            final_breakdown = b_taxi
            final_cost = taxi_sum
            final_type = "taxi"

    if is_ai_fare:
        recommend_msg += " / AI相場検索(運賃1.3倍マージン)"

    if DEBUG_MODE:
        with st.expander("🔍 [DEBUG] 費用計算の内訳", expanded=False):
            st.markdown(f"**mode:** `{selected_mode}` | **time:** {time_min}分 ({travel_hours:.1f}h)")
            st.markdown(f"**nights:** {nights} | **taxi_km:** {base_taxi_km:.1f}")
            if selected_mode == "flight":
                st.markdown(f"**flight_fare(片道flex):** {flight_fare:,} | **access:** {access_train_fare:,}")
            else:
                st.markdown(f"**total_fare(片道):** {navitime_route['total_fare']:,}")
            if not no_taxi:
                st.markdown("**タクシー:** " + str(b_taxi) + f" = {taxi_sum:,}")
            if not no_rental:
                st.markdown("**レンタカー:** " + str(b_rental) + f" = {rental_sum:,}")
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
# Streamlit UI
# ============================================================

st.set_page_config(page_title="交通費・出張見積もりアプリ", page_icon="🚗", layout="wide")
st.title("🚗 交通費・出張見積もりアプリ")

if DEBUG_MODE:
    st.caption("🐛 デバッグモード ON")

gemini_api_key, navitime_api_key = get_api_keys()

col1, col2 = st.columns([2, 1])
with col1:
    address_input = st.text_input("目的地（住所や施設名）", "")
with col2:
    station_choice = st.selectbox("出発拠点", ["淀屋橋駅", "大宮駅", "両方比較"], index=0)

col_a, col_b = st.columns(2)
with col_a:
    headcount = st.number_input("作業人数（人）", min_value=1, value=2)
with col_b:
    work_days = st.number_input("現地作業日数（日）", min_value=1, value=2)

st.markdown("**🚫 使用しない交通手段:**")
col_c1, col_c2, col_c3, col_c4 = st.columns(4)
with col_c1:
    no_flight = st.checkbox("飛行機を使用しない", value=False)
with col_c2:
    no_shinkansen = st.checkbox("新幹線を使用しない", value=False)
with col_c3:
    no_rental = st.checkbox("
