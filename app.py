
# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（Streamlit Cloud公開版）
運賃計算ルール:
  - 在来線のみ → unit_0（乗車券）のみ
  - 新幹線あり → unit_0 + unit_3（指定席特急券）
  - 飛行機 → flex（普通運賃）
  - グリーン車料金は常に除外
"""

import json
import math
import io
import re
import time
import requests
from datetime import datetime, timedelta
import streamlit as st

from google import genai
from google.genai import types

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 定数
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
            "Streamlit Cloud の Settings > Secrets に以下の形式で登録してください:\n\n"
            "```toml\n[api_keys]\ngemini = \"YOUR_KEY\"\nnavitime = \"YOUR_KEY\"\n```"
        )
        st.stop()
        return "", ""


# ============================================================
# Gemini リトライラッパー
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
                    f"サーバー混雑中... {wait_sec}秒後にリトライ（{attempt+1}/{MAX_RETRIES}）"
                )
            time.sleep(wait_sec)
            if retry_status_placeholder:
                retry_status_placeholder.empty()
    raise last_error


# ============================================================
# ユーティリティ
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
# Excel出力（社内フォーマット完全再現）
# ============================================================

def create_excel_report(pattern_data, address, headcount, work_days, origin_station):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交通費"

    # スタイル
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

    # 列幅
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

    # データ取得
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

    train_from, train_to = parse_route(routes.get("train", "-"))
    flight_from, flight_to = parse_route(routes.get("flight", "-"))
    access_from, access_to = parse_route(routes.get("access", "-"))

    def get_amount(key_contains):
        for k, v in breakdown.items():
            if key_contains in k:
                return int(v)
        return 0

    train_total = get_amount("電車・新幹線運賃")
    flight_total = get_amount("航空券費用")
    access_total = get_amount("アクセス電車運賃")
    rental_total = get_amount("レンタカー費用")
    taxi_total = get_amount("タクシー運賃")
    hotel_total = get_amount("宿泊費")

    is_train = train_total > 0
    is_flight = flight_total > 0
    is_access = access_total > 0
    is_rental = rental_total > 0
    is_taxi = taxi_total > 0
    is_hotel = hotel_total > 0

    # N列: 往復1人分
    train_n = int(train_total / headcount) if headcount > 0 and is_train else 0
    flight_n = int(flight_total / headcount) if headcount > 0 and is_flight else 0
    access_n = int(access_total / headcount) if headcount > 0 and is_access else 0
    rental_n = RENTAL_CAR_COST_PER_DAY if is_rental else 0
    taxi_n = int(taxi_total) if is_taxi else 0
    hotel_n = HOTEL_COST_PER_NIGHT_PER_PERSON if is_hotel else 0

    nights = 0
    match_nights = re.search(r'(\d+)泊', stay_note)
    if match_nights:
        nights = int(match_nights.group(1))
    elif is_hotel:
        nights = work_days

    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    flight_origin_airport = "伊丹空港" if "淀屋橋" in origin_station else "羽田空港"

    # ===== 設置設定作業 (Row 1-15) =====
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

    # Row 8 ヘッダー
    hdr_map = {'B': 'チェック', 'C': '項目', 'H': '経路', 'M': '調整費用',
               'N': '実費用', 'O': '日数・泊数', 'P': '人数', 'Q': '合計', 'R': '片路移動時間（h）'}
    for col_letter, label in hdr_map.items():
        cell = ws[f'{col_letter}8']
        cell.value = label
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr
    for col in range(4, 8):
        c = ws.cell(row=8, column=col)
        c.font = font_hdr
        c.fill = hdr_fill
        c.alignment = align_center
        c.border = border_hdr
    for col in range(9, 13):
        c = ws.cell(row=8, column=col)
        c.font = font_hdr
        c.fill = hdr_fill
        c.alignment = align_center
        c.border = border_hdr

    # Row 9-14 データ行
    rows_def = [
        {"ck": "✓" if is_train else "-", "item": "電車・新幹線（往復）　",
         "h": origin_station if is_train else "", "i": "駅->",
         "j": train_to if is_train else "", "k": "駅",
         "n": train_n, "o": 1, "p": "=C4",
         "r": time_hours if is_train else "", "s": ""},
        {"ck": "✓" if is_flight else "-", "item": "飛行機（往復）　",
         "h": flight_from if is_flight else flight_origin_airport, "i": "空港->",
         "j": flight_to if is_flight else "", "k": "空港",
         "n": flight_n, "o": 1, "p": "=C4",
         "r": time_hours if is_flight and not is_train else "", "s": "ANA/フレックスで試算"},
        {"ck": "✓" if is_access else "-", "item": "電車・新幹線（往復）　",
         "h": access_from if is_access else "", "i": "駅->",
         "j": access_to if is_access else "", "k": "駅",
         "n": access_n, "o": 1, "p": "=C4", "r": "", "s": ""},
        {"ck": "✓" if is_rental else "-", "item": "レンタカー",
         "h": "", "i": "", "j": "", "k": "",
         "n": rental_n, "o": rental_days if is_rental else 1, "p": "1",
         "r": 1 if is_rental else "", "s": "使用基準決める"},
        {"ck": "✓" if is_taxi else "-", "item": "タクシー（往復）",
         "h": "", "i": "", "j": "", "k": "",
         "n": taxi_n, "o": "=C5+C6", "p": "1", "r": "", "s": "バス/使用基準決める"},
        {"ck": "✓" if is_hotel else "-", "item": "宿泊",
         "h": "", "i": "", "j": "", "k": "",
         "n": hotel_n, "o": nights if is_hotel else 1, "p": "=C4", "r": "-", "s": ""},
    ]

    for idx, rd in enumerate(rows_def):
        rn = 9 + idx
        ws.cell(rn, 2, rd["ck"]).font = font_normal
        ws.cell(rn, 2).alignment = align_center
        ws.cell(rn, 2).border = border_data
        if rd["ck"] == "✓":
            ws.cell(rn, 2).fill = yellow_fill

        ws.cell(rn, 3, rd["item"]).font = font_normal
        ws.cell(rn, 3).alignment = align_left
        ws.cell(rn, 3).border = border_data

        ws.cell(rn, 8, rd["h"]).font = font_normal
        ws.cell(rn, 8).border = border_data
        if rd["h"] and rd["ck"] == "✓":
            ws.cell(rn, 8).fill = yellow_fill

        ws.cell(rn, 9, rd["i"]).font = font_normal
        ws.cell(rn, 9).alignment = align_center
        ws.cell(rn, 9).border = border_data

        ws.cell(rn, 10, rd["j"]).font = font_normal
        ws.cell(rn, 10).border = border_data
        if rd["j"] and rd["ck"] == "✓":
            ws.cell(rn, 10).fill = yellow_fill

        ws.cell(rn, 11, rd["k"]).font = font_normal
        ws.cell(rn, 11).alignment = align_center
        ws.cell(rn, 11).border = border_data

        # M列
        if rd["item"] in ("レンタカー", "宿泊"):
            ws.cell(rn, 13, rd["n"]).font = font_normal
        else:
            ws.cell(rn, 13, f"=ROUNDUP(N{rn},-3)").font = font_normal
        ws.cell(rn, 13).alignment = align_center
        ws.cell(rn, 13).border = border_data
        ws.cell(rn, 13).number_format = num_fmt

        # N列
        ws.cell(rn, 14, rd["n"]).font = font_normal
        ws.cell(rn, 14).alignment = align_center
        ws.cell(rn, 14).border = border_data
        ws.cell(rn, 14).number_format = num_fmt
        if rd["ck"] == "✓":
            ws.cell(rn, 14).fill = yellow_fill

        # O列
        ws.cell(rn, 15, rd["o"]).font = font_normal
        ws.cell(rn, 15).alignment = align_center
        ws.cell(rn, 15).border = border_data
        if rd["ck"] == "✓" and isinstance(rd["o"], int):
            ws.cell(rn, 15).fill = yellow_fill

        # P列
        ws.cell(rn, 16, rd["p"]).font = font_normal
        ws.cell(rn, 16).alignment = align_center
        ws.cell(rn, 16).border = border_data

        # Q列
        q = f'=IF(B{rn}="✓",SUM(M{rn}*O{rn}*P{rn}),IF(B{rn}="-",0,"確認"))'
        ws.cell(rn, 17, q).font = font_normal
        ws.cell(rn, 17).alignment = align_center
        ws.cell(rn, 17).border = border_data
        ws.cell(rn, 17).number_format = num_fmt

        # R列
        ws.cell(rn, 18, rd["r"]).font = font_normal
        ws.cell(rn, 18).alignment = align_center
        ws.cell(rn, 18).border = border_data
        ws.cell(rn, 18).number_format = num_fmt
        if rd["r"] and rd["r"] != "-" and rd["ck"] == "✓":
            ws.cell(rn, 18).fill = yellow_fill

        # S列
        if rd["s"]:
            ws.cell(rn, 19, rd["s"]).font = font_normal

    # Row 15 合計
    ws.cell(15, 17, "=SUM(Q9:Q14)").font = font_bold
    ws.cell(15, 17).number_format = num_fmt
    ws.cell(15, 18, '=SUMIF(B9:B14,"✓",R9:R14)').font = font_bold
    ws.cell(15, 18).number_format = num_fmt

    # ===== 下見作業 (Row 17-31) =====
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

    # Row 24 ヘッダー
    ws.merge_cells('C24:G24')
    ws.merge_cells('H24:L24')
    for col_letter, label in hdr_map.items():
        cell = ws[f'{col_letter}24']
        cell.value = label
        cell.font = font_hdr
        cell.fill = hdr_fill
        cell.alignment = align_center
        cell.border = border_hdr
    for col in range(4, 8):
        c = ws.cell(row=24, column=col)
        c.font = font_hdr
        c.fill = hdr_fill
        c.alignment = align_center
        c.border = border_hdr
    for col in range(9, 13):
        c = ws.cell(row=24, column=col)
        c.font = font_hdr
        c.fill = hdr_fill
        c.alignment = align_center
        c.border = border_hdr

    # Row 25-30 下見データ
    shitami = [
        {"ck": "-", "item": "電車・新幹線（往復）　", "h": "=H9", "i": "駅->", "j": "=J9", "k": "駅",
         "m": "=ROUNDUP(N25,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R9"},
        {"ck": "-", "item": "飛行機（往復）　", "h": flight_origin_airport, "i": "空港->", "j": "", "k": "空港",
         "m": "=ROUNDUP(N26,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R10"},
        {"ck": "-", "item": "電車・新幹線（往復）　", "h": "", "i": "駅->", "j": "", "k": "駅",
         "m": "=ROUNDUP(N27,-3)", "n": "", "o": 1, "p": "=C20", "r": "=R11"},
        {"ck": "-", "item": "レンタカー", "h": "", "i": "", "j": "", "k": "",
         "m": 12000, "n": "", "o": "=C21+C22", "p": "1", "r": "=R12"},
        {"ck": "-", "item": "タクシー（往復）", "h": "", "i": "", "j": "", "k": "",
         "m": "=ROUNDUP(N29,-3)", "n": "", "o": "=C21+C22", "p": "1", "r": "=R13"},
        {"ck": "-", "item": "宿泊", "h": "", "i": "", "j": "", "k": "",
         "m": 20000, "n": 20000, "o": "=C21+C22-1", "p": "=C20", "r": "-"},
    ]

    for idx, rd in enumerate(shitami):
        rn = 25 + idx
        ws.cell(rn, 2, rd["ck"]).font = font_normal
        ws.cell(rn, 2).alignment = align_center
        ws.cell(rn, 2).border = border_data
        ws.cell(rn, 3, rd["item"]).font = font_normal
        ws.cell(rn, 3).alignment = align_left
        ws.cell(rn, 3).border = border_data
        ws.cell(rn, 8, rd["h"]).font = font_normal
        ws.cell(rn, 8).border = border_data
        ws.cell(rn, 9, rd["i"]).font = font_normal
        ws.cell(rn, 9).alignment = align_center
        ws.cell(rn, 9).border = border_data
        ws.cell(rn, 10, rd["j"]).font = font_normal
        ws.cell(rn, 10).border = border_data
        ws.cell(rn, 11, rd["k"]).font = font_normal
        ws.cell(rn, 11).alignment = align_center
        ws.cell(rn, 11).border = border_data
        ws.cell(rn, 13, rd["m"]).font = font_normal
        ws.cell(rn, 13).alignment = align_center
        ws.cell(rn, 13).border = border_data
        ws.cell(rn, 13).number_format = num_fmt
        if rd["n"] != "":
            ws.cell(rn, 14, rd["n"]).font = font_normal
            ws.cell(rn, 14).border = border_data
            ws.cell(rn, 14).number_format = num_fmt
        ws.cell(rn, 15, rd["o"]).font = font_normal
        ws.cell(rn, 15).alignment = align_center
        ws.cell(rn, 15).border = border_data
        ws.cell(rn, 16, rd["p"]).font = font_normal
        ws.cell(rn, 16).alignment = align_center
        ws.cell(rn, 16).border = border_data
        q = f'=IF(B{rn}="✓",SUM(M{rn}*O{rn}*P{rn}),IF(B{rn}="-",0,"確認"))'
        ws.cell(rn, 17, q).font = font_normal
        ws.cell(rn, 17).alignment = align_center
        ws.cell(rn, 17).border = border_data
        ws.cell(rn, 17).number_format = num_fmt
        ws.cell(rn, 18, rd["r"]).font = font_normal
        ws.cell(rn, 18).alignment = align_center
        ws.cell(rn, 18).border = border_data
        ws.cell(rn, 18).number_format = num_fmt

    # Row 31 合計
    ws.cell(31, 17, "=SUM(Q25:Q30)").font = font_bold
    ws.cell(31, 17).number_format = num_fmt
    ws.cell(31, 18, '=SUMIF(B25:B30,"✓",R25:R30)').font = font_bold
    ws.cell(31, 18).number_format = num_fmt

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini 目的地分析
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
    prompt = (
        "以下の住所・施設名について、基本情報をJSON形式で回答してください。\n"
        "Google検索は不要です。あなたの知識のみで回答してください。\n\n"
        f"【対象】: {raw_address}\n\n"
        "以下の形式のJSONテキストのみを回答してください。\n"
        '{\n'
        '  "normalized_address": "対象の正式な住所（都道府県から）",\n'
        '  "dest_lat": 目的地の緯度(数値),\n'
        '  "dest_lon": 目的地の経度(数値),\n'
        '  "is_island_or_remote": 離島ならtrue/本土はfalse,\n'
        '  "nearest_airport_name": "離島の場合のみ最寄り空港名"\n'
        '}'
    )
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
            parsed = parse_json_from_text(text)
            if parsed:
                fallback_data.update(parsed)
    except Exception as e:
        st.warning(f"Gemini API失敗（フォールバック使用）: {e}")
    finally:
        retry_placeholder.empty()

    if not fallback_data.get("dest_lat") or not fallback_data.get("dest_lon"):
        lat, lon = geocode_fallback(fallback_data["normalized_address"])
        fallback_data["dest_lat"] = lat
        fallback_data["dest_lon"] = lon

    if fallback_data.get("is_island_or_remote"):
        guessed = guess_airport_from_address(fallback_data["normalized_address"])
        if not fallback_data.get("nearest_airport_name") or fallback_data["nearest_airport_name"] == "最寄り空港":
            fallback_data["nearest_airport_name"] = guessed

    return fallback_data


# ============================================================
# STEP 1.5: 離島 → Gemini + Google Search
# ============================================================

def search_flight_fare_with_gemini(raw_address, airport_name, gemini_key, origin_name):
    client = genai.Client(api_key=gemini_key.strip())
    origin_city = "大阪" if "淀屋橋" in origin_name else "東京"
    prompt = (
        "出張旅費算出のためGoogle検索で調査してください。\n\n"
        f"【出発地】: {origin_city}\n"
        f"【到着空港】: {airport_name}\n"
        f"【最終目的地】: {raw_address}\n\n"
        "以下の形式のJSONテキストのみを回答してください。\n"
        '{\n'
        '  "airport_lat": 到着空港の緯度,\n'
        '  "airport_lon": 到着空港の経度,\n'
        f'  "flight_fare_estimate": {origin_city}から{airport_name}への片道普通運賃(円),\n'
        f'  "flight_time_min": {origin_city}から{airport_name}までの片道所要時間(分)\n'
        '}'
    )
    fallback_data = {"airport_lat": None, "airport_lon": None,
                     "flight_fare_estimate": 60000, "flight_time_min": 180}
    retry_placeholder = st.empty()
    try:
        tool = types.Tool(google_search=types.GoogleSearch())
        config = types.GenerateContentConfig(tools=[tool], temperature=0.1)
        text = call_gemini_with_retry(client=client, model=GEMINI_MODEL, contents=prompt,
                                      config=config, retry_status_placeholder=retry_placeholder)
        if text:
            parsed = parse_json_from_text(text)
            if parsed:
                fallback_data.update(parsed)
    except Exception as e:
        st.warning(f"運賃検索失敗（フォールバック使用）: {e}")
    finally:
        retry_placeholder.empty()
    return fallback_data


# ============================================================
# STEP 2: NAVITIME API
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, navitime_key,
                               no_flight=False, no_shinkansen=False):
    headers = {"X-RapidAPI-Key": navitime_key.strip(), "X-RapidAPI-Host": NAVITIME_HOST}
    future_day = datetime.now() + timedelta(days=21)
    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "start_time": future_day.strftime("%Y-%m-%dT09:00:00"),
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
            st.warning(f"NAVITIME エラー: HTTP {res.status_code}")
            return None
        data = res.json()
        items = data.get("items", [])
        if not items:
            st.warning("NAVITIME: ルートが見つかりません。")
            return None

        fastest = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
        time_min = fastest.get("summary", {}).get("move", {}).get("time", 0)
        move_info = fastest.get("summary", {}).get("move", {})

        move_types = move_info.get("move_type", [])
        has_superexpress = "superexpress_train" in move_types

        fare_dict = move_info.get("fare", {})
        fare_unit_0 = fare_dict.get("unit_0", 0) if isinstance(fare_dict, dict) else 0
        fare_unit_3 = fare_dict.get("unit_3", 0) if isinstance(fare_dict, dict) else 0

        total_fare = int(fare_unit_0 + fare_unit_3) if has_superexpress else int(fare_unit_0)

        has_flight = False
        flight_fare = 0
        start_station_name = None
        end_station_name = None
        end_station_lat = None
        end_station_lon = None
        flight_start_name = None
        flight_end_name = None

        sections = fastest.get("sections", [])
        for i, sec in enumerate(sections):
            m_type = sec.get("move", "")
            sec_type = sec.get("type", "")

            if sec_type == "point":
                if not start_station_name:
                    start_station_name = sec.get("name")
                if "goal" not in sec.get("name", "").lower():
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
            "has_flight": has_flight, "time_min": time_min,
            "total_fare": total_fare, "flight_fare": int(flight_fare),
            "access_train_fare": int(access_train_fare),
            "last_walk_min": last_walk_min,
            "start_station": start_station_name or "出発駅",
            "end_station": end_station_name or "到着駅",
            "end_station_lat": end_station_lat, "end_station_lon": end_station_lon,
            "flight_start": flight_start_name, "flight_end": flight_end_name
        }

        if DEBUG_MODE:
            with st.expander("🔍 [DEBUG] NAVITIME詳細", expanded=False):
                st.json(params)
                st.json(move_info)
                st.markdown(f"**move_type:** `{move_types}` | **superexpress:** `{has_superexpress}`")
                st.markdown(f"**運賃:** {total_fare:,} 円")
                st.json(result)
        return result
    except Exception as e:
        st.warning(f"NAVITIME エラー: {e}")
        return None


# ============================================================
# STEP 3: パターン生成
# ============================================================

def build_best_route_patterns(current_st_name, ai_info, navitime_route,
                              headcount, work_days, no_rental=False, no_taxi=False):
    patterns = []

    if navitime_route:
        has_flight = navitime_route["has_flight"]
        time_min = navitime_route["time_min"]
        dest_lat, dest_lon = ai_info.get("dest_lat"), ai_info.get("dest_lon")
        e_lat, e_lon = navitime_route.get("end_station_lat"), navitime_route.get("end_station_lon")

        if e_lat and e_lon and dest_lat and dest_lon:
            base_taxi_km = haversine_km(e_lat, e_lon, dest_lat, dest_lon) * 1.3
            base_taxi_km = 0.0 if base_taxi_km <= 1.2 else max(base_taxi_km, 1.5)
        else:
            base_taxi_km = 0.0 if navitime_route["last_walk_min"] <= TAXI_WALK_THRESHOLD_MIN else max(navitime_route["last_walk_min"] * 0.08, 1.5)

        if has_flight:
            selected_mode = "flight"
            flight_fare = navitime_route["flight_fare"]
            access_train_fare = navitime_route["access_train_fare"]
            origin_airport = navitime_route["flight_start"] or "出発空港"
            airport_name = navitime_route["flight_end"] or "到着空港"
            end_st = navitime_route["end_station"]
            access_str = (f"{current_st_name} ➔ {origin_airport}" if airport_name == end_st
                          else f"{current_st_name} ➔ {origin_airport} / {airport_name} ➔ {end_st}")
            display_route_str = f"{current_st_name} ➔ {origin_airport} ➔ {airport_name} ➔ {end_st} ➔ 目的地"
            route_dict = {"train": "-", "flight": f"{origin_airport} ➔ {airport_name}",
                          "access": access_str,
                          "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "徒歩",
                          "rental": f"{end_st} ➔ 目的地周辺"}
            route_title = f"✈️ 飛行機ルート ({airport_name}経由)"
            is_ai_fare = False
        else:
            selected_mode = "train"
            end_st = navitime_route["end_station"]
            route_dict = {"train": f"{current_st_name} ➔ {end_st}", "flight": "-", "access": "-",
                          "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "徒歩",
                          "rental": f"{end_st} ➔ 目的地周辺"}
            route_title = f"🚄 電車ルート ({end_st}着)"
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
        route_dict = {"train": "-", "flight": f"{origin_airport} ➔ {airport_name}",
                      "access": f"{current_st_name} ➔ {origin_airport}",
                      "taxi": f"{airport_name} ↔ 目的地", "rental": f"{airport_name} ➔ 目的地周辺"}
        base_taxi_km = airport_to_dest_km
        route_title = f"✈️ 飛行機ルート ({airport_name}利用)"
        display_route_str = f"{current_st_name} ➔ {origin_airport} ➔ {airport_name} ➔ 目的地"
        is_ai_fare = True

    # 宿泊判定
    travel_hours = time_min / 60.0
    if selected_mode == "flight" or travel_hours >= 4.0:
        nights = work_days + 1
        stay_note = "前泊/後泊想定（" + str(nights) + "泊）"
    elif travel_hours >= 2.5:
        nights = work_days
        stay_note = "宿泊想定（" + str(nights) + "泊）"
    else:
        nights = max(work_days - 1, 0)
        stay_note = "標準（" + str(nights) + "泊）"

    # 費用計算
    hotel_cost = round_up_1000(HOTEL_COST_PER_NIGHT_PER_PERSON * headcount * nights)
    rental_days = work_days + (1 if "前泊" in stay_note else 0)
    rental_car_total = round_up_1000(RENTAL_CAR_COST_PER_DAY * rental_days)
    taxi_one_way = base_taxi_km * TAXI_FARE_PER_KM
    taxi_trips = nights + 1
    taxi_total = round_up_1000(taxi_one_way * 2 * taxi_trips)

    base_transport = {}
    if selected_mode == "flight":
        base_transport["航空券費用(往復・人数分・flex)"] = round_up_1000(flight_fare * 2 * headcount)
        base_transport["アクセス電車運賃(往復・人数分)"] = round_up_1000(access_train_fare * 2 * headcount)
    else:
        base_transport["電車・新幹線運賃(往復・人数分)"] = round_up_1000(navitime_route["total_fare"] * 2 * headcount)

    b_taxi = dict(base_transport)
    if taxi_total > 0:
        b_taxi["現地タクシー運賃(往復x" + str(taxi_trips) + "回)"] = taxi_total
    b_taxi["宿泊費"] = hotel_cost
    taxi_sum = sum(b_taxi.values())

    b_rental = dict(base_transport)
    b_rental["レンタカー費用(12000円x" + str(rental_days) + "日)"] = rental_car_total
    b_rental["宿泊費"] = hotel_cost
    rental_sum = sum(b_rental.values())

    if no_taxi and no_rental:
        b_walk = dict(base_transport)
        b_walk["宿泊費"] = hotel_cost
        final_breakdown = b_walk
        final_cost = sum(b_walk.values())
        final_type = "walk"
        final_name = route_title + " (徒歩)"
        recommend_msg = "最安ルート (徒歩前提)"
    elif no_rental:
        final_breakdown = b_taxi
        final_cost = taxi_sum
        final_type = "taxi"
        final_name = route_title + (" (徒歩)" if taxi_total == 0 else " + タクシー")
        recommend_msg = "最安ルート (タクシー)"
    elif no_taxi:
        final_breakdown = b_rental
        final_cost = rental_sum
        final_type = "rental"
        final_name = route_title + " + レンタカー"
        recommend_msg = "最安ルート (レンタカー)"
    else:
        if rental_sum < taxi_sum:
            final_breakdown = b_rental
            final_cost = rental_sum
            final_type = "rental"
            final_name = route_title + " + レンタカー"
            recommend_msg = "最安ルート (レンタカー・タクシーより " + f"{taxi_sum - rental_sum:,}" + " 円お得)"
        else:
            final_breakdown = b_taxi
            final_cost = taxi_sum
            final_type = "taxi"
            final_name = route_title + (" (徒歩)" if taxi_total == 0 else " + タクシー")
            recommend_msg = "最安ルート (タクシー)"

    if is_ai_fare:
        recommend_msg += " / AI相場(1.3倍マージン)"

    patterns.append({
        "type": final_type, "name": final_name, "time_min": time_min,
        "cost": final_cost, "breakdown": final_breakdown, "note": stay_note,
        "routes": route_dict, "display_route": display_route_str,
        "recommend_reason": recommend_msg, "is_recommended": True
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

st.markdown("**使用しない交通手段:**")
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

if st.button("最速出張見積もりを計算する", type="primary"):
    if not address_input.strip():
        st.warning("目的地を入力してください。")
        st.stop()

    if no_flight and no_shinkansen:
        st.info("飛行機・新幹線の両方を除外。在来線のみで検索します。")

    stations = (STATION_COORDS if station_choice == "両方比較"
                else {station_choice: STATION_COORDS[station_choice]})

    for current_st_name, (current_lat, current_lon) in stations.items():
        st.markdown(f"### 出発地: 【{current_st_name}】")

        with st.spinner(f"AIが {address_input} を分析中..."):
            ai_info = analyze_destination_with_gemini(address_input, gemini_api_key, current_st_name)

        if not ai_info.get("dest_lat") or not ai_info.get("dest_lon"):
            st.error("座標取得失敗。住所を詳しく入力してください。")
            continue

        st.success(f"検索地点: {ai_info.get('normalized_address', address_input)}")

        if DEBUG_MODE:
            with st.expander("🔍 [DEBUG] Gemini結果", expanded=False):
                st.json(ai_info)

        train_route = None
        if ai_info.get("is_island_or_remote", False) and not no_flight:
            st.info("離島判定 → 飛行機ルートを適用")
            airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
            with st.spinner(f"{airport_name} への運賃を検索中..."):
                fare_info = search_flight_fare_with_gemini(
                    address_input, airport_name, gemini_api_key, current_st_name)
            ai_info.update(fare_info)
            if not ai_info.get("airport_lat"):
                ai_info["airport_lat"] = ai_info.get("dest_lat", 24.3964)
                ai_info["airport_lon"] = ai_info.get("dest_lon", 124.2450)
        elif ai_info.get("is_island_or_remote", False) and no_flight:
            st.warning("離島ですが飛行機除外。船便等を検索します。")
            with st.spinner("NAVITIMEで検索中..."):
                train_route = get_navitime_fastest_route(
                    current_lat, current_lon, ai_info["dest_lat"], ai_info["dest_lon"],
                    navitime_api_key, no_flight=no_flight, no_shinkansen=no_shinkansen)
        else:
            with st.spinner("NAVITIMEで最速ルートを検索中..."):
                train_route = get_navitime_fastest_route(
                    current_lat, current_lon, ai_info["dest_lat"], ai_info["dest_lon"],
                    navitime_api_key, no_flight=no_flight, no_shinkansen=no_shinkansen)

        patterns = build_best_route_patterns(
            current_st_name, ai_info, train_route, headcount, work_days,
            no_rental=no_rental, no_taxi=no_taxi)

        for p in patterns:
            with st.expander(f"{p['name']} — 合計 {p['cost']:,} 円", expanded=p["is_recommended"]):
                st.success(p['recommend_reason'])
                st.write(f"**ルート:** {p['display_route']}")
                st.write(f"**片道:** 約 {p['time_min']} 分")
                st.write(f"**宿泊:** {p['note']}")
                st.write("**内訳:**")
                for item, amt in p["breakdown"].items():
                    st.write(f"　・{item}: **{amt:,}** 円")

                excel_bytes = create_excel_report(
                    p, ai_info.get('normalized_address', address_input),
                    headcount, work_days, current_st_name)
                st.download_button(
                    label="Excel見積書をダウンロード",
                    data=excel_bytes,
                    file_name=f"交通費見積_{current_st_name}発_{address_input[:10]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{current_st_name}_{p['type']}")
        st.markdown("---")

