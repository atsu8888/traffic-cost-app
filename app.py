# -*- coding: utf-8 -*-
"""
交通費・出張見積もりアプリ（指定Excelフォーマット完全準拠版）

機能:
  - 離島/遠方はAIで判定し、NAVITIMEをスキップしてGoogle検索の運賃(1.3倍)を採用
  - NAVITIMEの「goal」表記を確実に除外し、最終駅(例: 南新庄駅)を抽出
  - ユーザー指定のExcelフォーマット（A1〜H14等の固定レイアウト）に完全に準拠
  - 現地移動（タクシー vs レンタカー）を内部で比較し、安い方「だけ」を1つ出力
  - 試算結果を社内フォーマットのExcelファイル(.xlsx)として自動出力
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
from openpyxl.styles import Font, Alignment, Border, Side

# ============================================================
# 定数・固定値
# ============================================================

STATION_COORDS = {
    "大宮駅": (35.906702, 139.623596),
    "淀屋橋駅": (34.693729, 135.499814),
}

HOTEL_COST_PER_NIGHT_PER_PERSON = 20_000   # 宿泊費（1泊/1名）
RENTAL_CAR_COST_PER_DAY = 12_000           # レンタカー（1日）
TAXI_FARE_PER_KM = 400                    # タクシー単価（円/km）
TAXI_WALK_THRESHOLD_MIN = 15              # 徒歩15分以内ならタクシー不要

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
# Excelファイル出力関数（指定フォーマット完全準拠）
# ============================================================

def create_excel_report(pattern_data: dict, address: str, headcount: int, work_days: int) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交通費"
    
    font_normal = Font(name="ＭＳ Ｐゴシック", size=11)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # 基本情報の書き込み
    ws["B2"] = "設置設定作業"
    ws["H2"] = "基本大阪と埼玉のどちらか近い方から算出するが、TIS、CSI案件のみ全て大阪から算出する。"
    ws["S2"] = "作業日数"
    ws["T2"] = "サーバ台数"
    
    ws["H3"] = "大阪：淀屋橋駅、埼玉：大宮駅"
    ws["S3"] = "2日（作業日1日、予備日：1日）"
    ws["T3"] = "1台"
    
    ws["B4"] = "住所"
    ws["C4"] = address  # 住所はフォーマットだと特に指定行がないためB4の横に配置と仮定
    ws["S4"] = "4日（作業日3日、予備日：1日）"
    ws["T4"] = "2～5台まで"
    
    ws["B5"] = "人数"
    ws["C5"] = headcount
    ws["D5"] = "人"
    ws["S5"] = "5日（作業日4日、予備日：1日）"
    ws["T5"] = "6台以上"
    
    ws["B6"] = "作業"
    ws["C6"] = work_days
    ws["D6"] = "日"
    
    travel_days = 2 if "前泊" in pattern_data.get("note", "") else 1
    ws["B7"] = "移動"
    ws["C7"] = travel_days
    ws["D7"] = "日"
    ws["E7"] = "※移動に2.5h以上かかる場合は宿泊想定／移動に4h以上かかる場合又は飛行機を利用の場合はさらに前泊か後泊想定"

    ws["M8"] = "※最大費用の経路を想定"

    # ヘッダー行
    headers = {
        "B9": "チェック", "C9": "項目", "H9": "経路", "M9": "調整費用", 
        "N9": "実費用", "O9": "日数・泊数", "P9": "人数", "Q9": "合計", "R9": "片路移動時間（h）"
    }
    for cell_pos, val in headers.items():
        ws[cell_pos] = val

    breakdown = pattern_data.get("breakdown", {})
    routes = pattern_data.get("routes", {})
    route_details = pattern_data.get("route_details", {}) # {start: 'xx', flight_start: 'xx', flight_end: 'xx', end: 'xx'}
    
    is_flight = "flight_start" in route_details and route_details["flight_start"]
    start_st = route_details.get("start", "出発駅")
    end_st = route_details.get("end", "到着駅")
    f_start = route_details.get("flight_start", "")
    f_end = route_details.get("flight_end", "")

    # 各行のデータ構築 (フォーマットに合わせたセル分割)
    # 電車・新幹線（アクセス/純粋）
    has_train = "電車・新幹線運賃" in list(breakdown.keys())[0] if breakdown else False
    if is_flight:
        train_val = breakdown.get("最寄り空港アクセス電車運賃(往復・人数分)", 0)
        train_chk = "✓" if train_val > 0 else "-"
        train_route_1 = start_st
        train_route_2 = f_start
    else:
        train_val = list(breakdown.values())[0] if breakdown else 0
        train_chk = "✓" if train_val > 0 else "-"
        train_route_1 = start_st
        train_route_2 = end_st

    # 飛行機
    flight_val = breakdown.get("航空券費用(往復・人数分)", 0) if is_flight else 0
    flight_chk = "-" if not is_flight else "✓"
    flight_route_1 = f_start if is_flight else ""
    flight_route_2 = f_end if is_flight else ""

    # レンタカー
    rental_key = next((k for k in breakdown if "レンタカー" in k), None)
    rental_val = breakdown[rental_key] if rental_key else 0
    rental_chk = "✓" if rental_val > 0 else "-"
    
    # タクシー
    taxi_key = next((k for k in breakdown if "タクシー" in k), None)
    taxi_val = breakdown[taxi_key] if taxi_key else 0
    taxi_chk = "✓" if taxi_val > 0 else "-"

    # 宿泊
    hotel_val = breakdown.get("宿泊費", 0)
    hotel_chk = "✓" if hotel_val > 0 else "-"

    # --- 行ごとの書き込み ---
    
    # Row 10: 電車・新幹線（往復） 又は アクセス電車（往復）出発側
    ws["B10"] = train_chk
    ws["C10"] = "電車・新幹線（往復）　"
    ws["H10"] = train_route_1
    ws["I10"] = "駅->"
    ws["J10"] = train_route_2
    ws["K10"] = "駅"
    ws["M10"] = train_val
    ws["N10"] = train_val
    ws["O10"] = 1
    ws["P10"] = headcount
    ws["Q10"] = train_val
    ws["R10"] = round(pattern_data.get("time_min", 0) / 60.0, 1)

    # Row 11: 飛行機（往復）
    ws["B11"] = flight_chk
    ws["C11"] = "飛行機（往復）　"
    ws["H11"] = flight_route_1
    ws["I11"] = "空港->" if is_flight else ""
    ws["J11"] = flight_route_2
    ws["K11"] = "空港" if is_flight else ""
    ws["M11"] = flight_val
    ws["N11"] = flight_val
    ws["O11"] = 1
    ws["P11"] = headcount
    ws["Q11"] = flight_val
    if is_flight and pattern_data.get("is_ai_fare"):
        ws["S11"] = "ANA/フレックスで試算(AI概算1.3倍)"

    # Row 12: 到着側のアクセス電車（必要な場合のみ）
    if is_flight and f_end != end_st:
        ws["B12"] = "✓"
        ws["C12"] = "電車・新幹線（往復）　"
        ws["H12"] = f_end
        ws["I12"] = "駅->"
        ws["J12"] = end_st
        ws["K12"] = "駅"
        # 料金はアクセス電車代（10行目）に合算されている想定とする
        ws["M12"] = 0
        ws["N12"] = 0
        ws["O12"] = 1
        ws["P12"] = headcount
        ws["Q12"] = 0
    else:
        ws["B12"] = "-"
        ws["C12"] = "電車・新幹線（往復）　"
        ws["I12"] = "駅->"
        ws["K12"] = "駅"

    # Row 13: レンタカー
    ws["B13"] = rental_chk
    ws["C13"] = "レンタカー"
    ws["M13"] = rental_val
    ws["O13"] = work_days + (1 if travel_days > 1 else 0)
    ws["P13"] = 1 # レンタカーは1台換算
    ws["Q13"] = rental_val

    # Row 14: タクシー
    ws["B14"] = taxi_chk
    ws["C14"] = "タクシー（往復）"
    ws["M14"] = taxi_val
    ws["O14"] = travel_days + work_days
    ws["P14"] = 1
    ws["Q14"] = taxi_val

    # Row 15: 宿泊
    ws["B15"] = hotel_chk
    ws["C15"] = "宿泊"
    ws["M15"] = HOTEL_COST_PER_NIGHT_PER_PERSON
    ws["O15"] = travel_days - 1 + work_days
    ws["P15"] = headcount
    ws["Q15"] = hotel_val

    # 合計行
    ws["Q16"] = pattern_data.get("cost", 0)
    ws["R16"] = round(pattern_data.get("time_min", 0) / 60.0, 1)

    # 全体フォント適用
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.font = font_normal

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STEP 1: Gemini による目的地分析 ＋ 国土地理院バックアップ
# ============================================================

def geocode_fallback(address: str):
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

以下の形式のJSONテキストのみを回答してください。
{{
  "normalized_address": "対象の正式な住所",
  "dest_lat": 目的地の緯度(数値),
  "dest_lon": 目的地の経度(数値),
  "is_island_or_remote": 対象が海を渡る完全な離島(沖縄、奄美、石垣など)ならtrue。北海道や本州等はfalse,
  "nearest_airport_name": "最寄り空港名(具体名必須。同じ県でも島が違う場合はその島の空港)",
  "airport_lat": 空港の緯度(数値),
  "airport_lon": 空港の経度(数値),
  "flight_fare_estimate": {origin_city}から最寄り空港への大人片道普通運賃概算(数値),
  "flight_time_min": {origin_city}から最寄り空港までの片道総所要時間・分(数値)
}}
"""
    model_name = get_active_gemini_model_name()
    
    islands = ["沖縄", "奄美", "沖永良部", "石垣", "宮古", "屋久島", "種子島", "徳之島", "与論", "久米島", "喜界"]
    is_island_guess = any(island in raw_address for island in islands)
    
    fallback_data = {
        "normalized_address": raw_address,
        "is_island_or_remote": is_island_guess,
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
# STEP 2: NAVITIME API (全国・全交通手段を検索)
# ============================================================

def get_navitime_fastest_route(start_lat, start_lon, goal_lat, goal_lon, api_key: str):
    clean_key = api_key.strip()
    headers = {"X-RapidAPI-Key": clean_key, "X-RapidAPI-Host": NAVITIME_HOST}
    
    future_day = datetime.now() + timedelta(days=21)
    start_time_iso = future_day.strftime("%Y-%m-%dT09:00:00")
    
    params = {
        "start": f"{start_lat},{start_lon}",
        "goal": f"{goal_lat},{goal_lon}",
        "start_time": start_time_iso,
        "format": "json"
    }

    try:
        res = requests.get(NAVITIME_URL, headers=headers, params=params, timeout=15)
        if res.status_code != 200: return None
        
        data = res.json()
        items = data.get("items", [])
        if not items: return None

        fastest_item = min(items, key=lambda x: x.get("summary", {}).get("move", {}).get("time", 99999))
        time_min = fastest_item.get("summary", {}).get("move", {}).get("time", 0)
        
        move_info = fastest_item.get("summary", {}).get("move", {})
        total_fare = move_info.get("fare", {}).get("unit_0", 0) if isinstance(move_info.get("fare"), dict) else 0

        has_flight = False
        flight_fare = 0
        
        start_station_name = None
        end_station_name = None
        end_station_lat = None
        end_station_lon = None
        flight_start_name = None
        flight_end_name = None

        sections = fastest_item.get("sections", [])
        
        for i, sec in enumerate(sections):
            m_type = sec.get("move", "")
            sec_type = sec.get("type", "")

            if sec_type == "point":
                p_name = sec.get("name", "").lower()
                if not start_station_name and "goal" not in p_name and "目的地" not in p_name:
                    start_station_name = sec.get("name")
                
                # goalを無視して駅名のみを記録
                if "goal" not in p_name and "目的地" not in p_name:
                    end_station_name = sec.get("name")
                    if "coord" in sec:
                        end_station_lat = sec["coord"].get("lat")
                        end_station_lon = sec["coord"].get("lon")
                
            if sec_type == "move" and m_type in ["plane", "flight", "air", "airplane", "aeroplane"]:
                has_flight = True
                if "fare" in sec and isinstance(sec["fare"], dict):
                    flight_fare += sec["fare"].get("unit_0", 0)
                    
                # 出発空港を記録
                if not flight_start_name:
                    for j in range(i-1, -1, -1):
                        if sections[j].get("type") == "point":
                            flight_start_name = sections[j].get("name")
                            break
                # 到着空港を記録
                for j in range(i+1, len(sections)):
                    if sections[j].get("type") == "point":
                        flight_end_name = sections[j].get("name")
                        break

        # 最後の移動手段にかかった時間
        last_walk_min = 0
        if sections and sections[-1].get("type") == "move":
            last_walk_min = sections[-1].get("time", 0)

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
            "last_walk_min": last_walk_min,
            "start_station": start_station_name or "出発駅",
            "end_station": end_station_name or "到着駅",
            "end_station_lat": end_station_lat,
            "end_station_lon": end_station_lon,
            "flight_start": flight_start_name,
            "flight_end": flight_end_name
        }
    except Exception as e:
        return None


# ============================================================
# パターン生成と比較ロジック（結果を一つに絞る）
# ============================================================

def build_best_route_patterns(st_name: str, ai_info: dict, navitime_route: dict, headcount: int, work_days: int):
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
                access_str = f"{st_name} ➔ {origin_airport}"
            else:
                access_str = f"{st_name} ➔ {origin_airport} / {airport_name} ➔ {end_st}"
                
            display_route_str = f"{st_name} ➔ {origin_airport} ➔ {airport_name} ➔ {end_st} ➔ 目的地"
            
            route_dict = {
                "train": "-",
                "flight": f"{origin_airport} ➔ {airport_name}",
                "access": access_str,
                "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "現地徒歩",
                "rental": f"{end_st} ➔ 目的地周辺"
            }
            route_details = {
                "start": st_name,
                "flight_start": origin_airport,
                "flight_end": airport_name,
                "end": end_st
            }
            route_title_prefix = f"✈️ 飛行機ルート ({airport_name}経由 {end_st}着)"
            is_ai_fare = False
            
        else:
            selected_mode = "train"
            end_st = navitime_route["end_station"]
            route_dict = {
                "train": f"{st_name} ➔ {end_st}",
                "flight": "-",
                "access": "-",
                "taxi": f"{end_st} ↔ 目的地" if base_taxi_km > 0 else "現地徒歩",
                "rental": f"{end_st} ➔ 目的地周辺"
            }
            route_details = {
                "start": st_name,
                "flight_start": "",
                "flight_end": "",
                "end": end_st
            }
            route_title_prefix = f"🚄 新幹線/電車ルート ({end_st}着)"
            display_route_str = f"{st_name} ➔ {end_st} ➔ 目的地"
            is_ai_fare = False
            
    else:
        # NAVITIMEをスキップした場合（石垣など完全離島）
        selected_mode = "flight"
        airport_name = ai_info.get("nearest_airport_name", "最寄り空港")
        airport_lat, airport_lon = ai_info.get("airport_lat", 0), ai_info.get("airport_lon", 0)
        dest_lat, dest_lon = ai_info.get("dest_lat", 0), ai_info.get("dest_lon", 0)
        
        airport_to_dest_km = 15.0
        if airport_lat and airport_lon and dest_lat and dest_lon:
            airport_to_dest_km = max(haversine_km(airport_lat, airport_lon, dest_lat, dest_lon) * 1.3, 1.0)
        
        time_min = int(60 + ai_info.get("flight_time_min", 150) + (airport_to_dest_km / 40.0 * 60))
        
        flight_fare = int(ai_info.get("flight_fare_estimate", 60000) * 1.3) 
        
        origin_airport = "伊丹空港" if "淀屋橋" in st_name else "羽田空港"
        access_train_fare = 1500 if "大宮" in st_name else 500
        
        route_dict = {
            "train": "-",
            "flight": f"{origin_airport} ➔ {airport_name}",
            "access": f"{st_name} ➔ {origin_airport}",
            "taxi": f"{airport_name} ↔ 目的地",
            "rental": f"{airport_name} ➔ 目的地周辺"
        }
        route_details = {
            "start": st_name,
            "flight_start": origin_airport,
            "flight_end": airport_name,
            "end": airport_name
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

    patterns.append({
        "type": final_type, 
        "name": final_name,
        "time_min": time_min, 
        "cost": final_cost,
        "breakdown": final_breakdown, 
        "note": stay_note,
        "routes": route_dict,
        "route_details": route_details,
        "display_route": display_route_str,
        "recommend_reason": recommend_msg,
        "is_ai_fare": is_ai_fare,
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
    address_input = st.text_input("目的地（住所や施設名）", "山形県新庄市若葉町12-1（新庄徳洲会病院）")
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
        
        train_route = None
        if ai_info.get("is_island_or_remote", False):
            st.info("💡 **海を渡る離島・遠隔地と判定されました。NAVITIME検索をスキップし、飛行機ルートを適用します。**")
        else:
            with st.spinner(f"🧭 NAVITIMEで全国対応の最速ルートを検索中..."):
                train_route = get_navitime_fastest_route(st_lat, st_lon, ai_info["dest_lat"], ai_info["dest_lon"], navitime_api_key)

        patterns = build_best_route_patterns(st_name, ai_info, train_route, headcount, work_days)

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


