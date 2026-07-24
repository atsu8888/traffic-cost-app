import io
import json
import openpyxl
import requests
import streamlit as st
import google.generativeai as genai
import re

# ===== 画面設定 =====
st.set_page_config(page_title="交通費見積自動生成ツール", page_icon="🚗", layout="wide")

st.title("🚗 訪問費・交通費 見積作成ツール")
st.caption("出発地・目的地・作業日数・宿泊数を入力し、AI解析結果入りの自社Excelを自動生成します")

# ===== サイドバー（APIキー設定＆モデル自動検出） =====
selected_model_name = None

with st.sidebar:
    st.header("🔑 初期設定")
    api_key = st.text_input("Gemini APIキーを入力", type="password", help="Google AI Studioで取得したAPIキーを入力してください")
    
    if api_key:
        try:
            genai.configure(api_key=api_key)
            available_models = [
                m.name for m in genai.list_models() 
                if 'generateContent' in m.supported_generation_methods
            ]
            
            if available_models:
                st.success("✅ APIキー接続成功")
                
                default_idx = 0
                for i, m_name in enumerate(available_models):
                    if "gemini-flash-latest" in m_name:
                        default_idx = i
                        break
                    elif "gemini-1.5-flash" in m_name and default_idx == 0:
                        default_idx = i
                
                selected_model_name = st.selectbox(
                    "使用するAIモデル（自動検出）", 
                    available_models, 
                    index=default_idx
                )
            else:
                st.error("❌ 利用可能なモデルが見つかりませんでした。")
        except Exception as e:
            st.error(f"❌ APIキー認証エラー: {e}")
            
    st.markdown("---")
    st.info("※ 入力したAPIキーは画面のセッション内のみで使用され、外部に保存・送信されません。")

# ===== メインフォーム =====
with st.form("estimation_form"):
    st.subheader("1. 移動区間（どこから ➔ どこまで）")
    col1, col2 = st.columns(2)
    with col1:
        origin_preset = st.selectbox("出発地", ["大宮駅（埼玉）", "淀屋橋駅（大阪）", "その他（直接入力）"])
        if origin_preset == "その他（直接入力）":
            origin_station = st.text_input("出発地を入力", value="東京駅")
        else:
            origin_station = origin_preset.split("（")[0]
            
    with col2:
        destination = st.text_input("目的地（住所または施設名）", value="〒730-8655 広島県広島市中区中島町3番30号（土谷総合病院）")

    st.subheader("2. 人数・日程設定（手動入力）")
    col3, col4, col5 = st.columns(3)
    with col3:
        people_count = st.number_input("人数", min_value=1, max_value=20, value=2)
    with col4:
        work_days = st.number_input("作業日数（何日）", min_value=1, max_value=30, value=4)
    with col5:
        stay_nights = st.number_input("宿泊数（何泊）", min_value=0, max_value=30, value=3)

    submitted = st.form_submit_button("✨ AIで条件解析 & Excel生成", use_container_width=True)

# ===== 試算＆Excel流し込み処理 =====
if submitted:
    if not api_key:
        st.error("⚠️ 左側のサイドバーに「Gemini APIキー」を入力してください！")
    elif not selected_model_name:
        st.error("⚠️ 有効なAIモデルが選択されていません。")
    else:
        with st.spinner(f"🤖 {selected_model_name.split('/')[-1]} が最適な経路・料金を計算＆Excel作成中..."):
            try:
                model = genai.GenerativeModel(selected_model_name)

                # --- AIへの指示（新幹線優先・飛行機判定3時間ルール・一番最寄りの駅ルール） ---
                prompt = f"""
                あなたはプロの営業事務として、見積作成のために交通費を試算します。
                以下の条件に基づき、移動ルートと概算金額を算出し、指定のJSONフォーマットのみで回答してください。
                ※絶対にJSON以外の挨拶や説明文を出力しないでください。

                【基本条件】
                ・出発地: {origin_station}
                ・目的地: {destination}
                ・人数: {people_count}名
                ・作業日数: {work_days}日
                ・宿泊数: {stay_nights}泊

                【移動ルート・交通機関の判定ルール】
                1. 【基本原則】できる限り新幹線や電車（在来線含む）を利用し、目的地に一番近い「最寄り駅」まで電車で移動することを最優先としてください。
                2. 【飛行機の採用条件】
                   ・原則は新幹線・電車ルートを優先します。
                   ・ただし、新幹線・電車ルートで移動する場合に比べ、飛行機（ANA国内線フレックス運賃）を利用した方が【3時間以上】所要時間を短縮できる場合に限り、飛行機ルートを採用してください（"main_transport_type": "flight"）。
                   ・短縮時間が3時間未満の場合は、新幹線・電車ルートを採用してください（"main_transport_type": "shinkansen_train"）。
                3. バスは使用しないでください。

                【最寄り駅からの現地移動手段（徒歩/タクシー/レンタカー）の判定ルール】
                ・最寄り駅から目的地まで「徒歩10分未満」の場合は【徒歩】（"selected_local_transport": "none"）としてください。
                ・最寄り駅（または最寄り空港）から目的地まで「徒歩10分以上」の場合は、以下の通り【タクシー】か【レンタカー】を比較し、経済的かつ合理的な方を選択してください。
                    * タクシーを使用する場合： 宿泊日数 + 1日分（往復）の費用で計算する。
                    * レンタカーを使用する場合： 1日12,000円 × (宿泊日数 + 1日分) で計算する。
                    * タクシー選択時は "taxi"、レンタカー選択時は "rentacar" に設定する。

                【出力フォーマット（これだけを出力）】
                {{
                    "origin_station": "{origin_station}",
                    "dest_station": "目的地に一番近い最寄り駅（または空港名）",
                    "main_transport_type": "shinkansen_train" または "flight",
                    "train_or_air_fare_one_way": 片道メイン交通機関(1名分)の概算運賃(整数),
                    "selected_local_transport": "taxi" または "rentacar" または "none",
                    "taxi_fare_one_way": タクシー選択時の片道運賃目安(整数。不要なら0),
                    "one_way_hours": 片道所要時間の数値(少数可),
                    "reason": "新幹線vs飛行機の時間差の解説、最寄り駅の指定理由、タクシー/レンタカーの選定理由"
                }}
                """

                # モデルの呼び出し
                response = model.generate_content(prompt)
                raw_text = response.text
                
                # JSON抽出
                match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                if match:
                    clean_text = match.group(0)
                else:
                    clean_text = raw_text

                try:
                    ai_result = json.loads(clean_text)
                except json.JSONDecodeError:
                    st.error("❌ AIの回答からデータを抽出できませんでした。以下はAIの実際の回答です：")
                    st.code(raw_text)
                    st.stop()

                # --- 費用計算 ---
                hotel_cost_per_night = 20000 
                rentacar_cost_per_day = 12000
                local_transport_days = stay_nights + 1  # 宿泊日数+1日分

                # 往復のメイン交通費
                main_round_trip_fare = ai_result.get('train_or_air_fare_one_way', 0) * 2

                # 現地交通費（タクシー vs レンタカー）の明細計算
                taxi_one_way = ai_result.get('taxi_fare_one_way', 0)
                taxi_total_fare = 0
                rentacar_total_fare = 0
                
                local_detail_str = ""
                if ai_result.get('selected_local_transport') == "taxi":
                    taxi_total_fare = (taxi_one_way * 2) * local_transport_days
                    local_detail_str = f"🚕 タクシー内訳: 片道 {taxi_one_way:,}円 × 往復 × {local_transport_days}日分 ＝ 合計 {taxi_total_fare:,}円"
                elif ai_result.get('selected_local_transport') == "rentacar":
                    rentacar_total_fare = rentacar_cost_per_day * local_transport_days
                    local_detail_str = f"🚗 レンタカー内訳: 1日 {rentacar_cost_per_day:,}円 × {local_transport_days}日分 ＝ 合計 {rentacar_total_fare:,}円"
                else:
                    local_detail_str = "🚶 徒歩移動 (最寄り駅から10分未満)"

                # 解析結果の画面表示
                display_name = selected_model_name.replace("models/", "")
                st.success(f"✅ AIによる経路・条件解析が完了しました！（使用モデル: {display_name}）")
                
                col_a, col_b, col_c, col_d = st.columns(4)
                
                transport_label = "✈️ 飛行機(ANA)" if ai_result.get('main_transport_type') == "flight" else "🚅 新幹線・電車"
                col_a.metric("メイン手段", transport_label)
                col_b.metric("最寄り駅/空港", ai_result.get('dest_station', '不明'))
                col_c.metric("片道運賃(1名)", f"{ai_result.get('train_or_air_fare_one_way', 0):,} 円")
                
                if ai_result.get('selected_local_transport') == "rentacar":
                    col_d.metric("現地手段", "🚗 レンタカー")
                elif ai_result.get('selected_local_transport') == "taxi":
                    col_d.metric("現地手段", "🚕 タクシー")
                else:
                    col_d.metric("現地手段", "🚶 徒歩(10分未満)")
                    
                st.caption(f"📝 AI判定理由: {ai_result.get('reason', '')}")
                st.info(f"💡 **【明細内訳】** {local_detail_str}")

                # -------------------------------------------------------------
                # 【Excel読み込み & 明細付き書き込み処理】
                # -------------------------------------------------------------
                excel_filename = "見積指標（CSI,SSI,TIS）_r1 - コピー 2.xlsx"
                wb = openpyxl.load_workbook(excel_filename)
                ws = wb["交通費"]

                # --- 基本情報の書き込み ---
                ws["C3"] = destination      
                ws["C4"] = people_count     
                ws["C5"] = work_days        

                dest_station_name = ai_result.get('dest_station', '')

                # --- Excelへの流し込み ---
                # メイン交通機関（行9:新幹線 or 行10:飛行機）
                if ai_result.get('main_transport_type') == "flight":
                    ws["B9"] = "-"
                    ws["H9"] = ""
                    ws["J9"] = ""
                    ws["N9"] = 0
                    
                    ws["B10"] = "✓"
                    ws["H10"] = ai_result['origin_station']
                    ws["J10"] = dest_station_name
                    ws["N10"] = main_round_trip_fare
                    ws["R10"] = ai_result['one_way_hours']
                else:
                    ws["B9"] = "✓"
                    ws["H9"] = ai_result['origin_station']
                    ws["J9"] = dest_station_name
                    ws["N9"] = main_round_trip_fare
                    ws["R9"] = ai_result['one_way_hours']
                    
                    ws["B10"] = "-"
                    ws["H10"] = ""
                    ws["J10"] = ""
                    ws["N10"] = 0

                ws["B11"] = "-" # 新幹線2

                # 行12: レンタカー（明細：発地に最寄り駅、着地に目的地を自動補完）
                if ai_result.get('selected_local_transport') == "rentacar":
                    ws["B12"] = "✓"
                    ws["H12"] = dest_station_name  # 発地: 最寄り駅/空港
                    ws["J12"] = destination         # 着地: 目的地
                    ws["N12"] = rentacar_total_fare
                else:
                    ws["B12"] = "-"
                    ws["H12"] = ""
                    ws["J12"] = ""
                    ws["N12"] = 0

                # 行13: タクシー（明細：発地に最寄り駅、着地に目的地を自動補完）
                if ai_result.get('selected_local_transport') == "taxi":
                    ws["B13"] = "✓"
                    ws["H13"] = dest_station_name  # 発地: 最寄り駅
                    ws["J13"] = destination         # 着地: 目的地
                    ws["N13"] = taxi_total_fare
                else:
                    ws["B13"] = "-"
                    ws["H13"] = ""
                    ws["J13"] = ""
                    ws["N13"] = 0

                # 行14: 宿泊
                if stay_nights > 0:
                    ws["B14"] = "✓"
                    ws["N14"] = hotel_cost_per_night  
                    ws["O14"] = stay_nights           
                else:
                    ws["B14"] = "-"
                    ws["N14"] = 0
                    ws["O14"] = 0

                # --- ファイル出力 ---
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                # ダウンロードボタン
                st.download_button(
                    label="📥 AI解析済みの見積Excelをダウンロード",
                    data=output,
                    file_name=f"交通費見積_{origin_station}発_{destination[:6]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")