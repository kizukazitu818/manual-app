import streamlit as st
import time
import os
import json
import datetime
import cv2
import re
import numpy as np
import google.generativeai as genai
from io import BytesIO
import base64
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
from gtts import gTTS
from google.generativeai.types import HarmCategory, HarmBlockThreshold
from streamlit_drawable_canvas import st_canvas
import streamlit_drawable_canvas as canvas_lib

# --- 0. 修正パッチ（お絵かき機能用） ---
def fix_canvas_library():
    # 画像をブラウザ用データに変換する関数を自作
    def custom_image_to_url(image, width, clamp, channels, output_format, image_id):
        try:
            buffered = BytesIO()
            image.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            return f"data:image/png;base64,{img_str}"
        except Exception:
            return ""
    
    # ライブラリに注入
    if hasattr(canvas_lib, 'st_image'):
        canvas_lib.st_image.image_to_url = custom_image_to_url

fix_canvas_library()

# --- 1. 基本設定 ---
st.set_page_config(page_title="Nano Factory AI", page_icon="📜", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=M+PLUS+Rounded+1c:wght@300;400;700&display=swap');
    html, body, [class*="css"] { font-family: 'M PLUS Rounded 1c', sans-serif !important; }
    [data-testid="stFileUploaderDropzone"] { background-color: #E6F3FF; border: 2px dashed #007BFF; border-radius: 15px; padding: 20px; }
    [data-testid="stSidebar"] { background-color: #E6F3FF; }
    h1 { border-bottom: 5px solid #FFD700; padding-bottom: 10px; }
    .step-card { border: 1px solid #ddd; padding: 15px; border-radius: 10px; margin-bottom: 10px; background-color: #f9f9f9; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 関数群 ---
@st.cache_data(ttl=600)
def get_available_models(api_key):
    default_models = ["gemini-1.5-flash", "gemini-2.0-flash-exp"]
    if not api_key: return default_models
    try:
        genai.configure(api_key=api_key)
        models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                if "deep-research" in name or "ultra" in name: continue
                models.append(name)
        models.sort()
        prioritized = [m for m in models if "flash" in m]
        others = [m for m in models if "flash" not in m]
        return prioritized + others if (prioritized + others) else default_models
    except: return default_models

def clean_timestamp(ts_value):
    if ts_value is None: return 0.0
    if isinstance(ts_value, (int, float)): return float(ts_value)
    s = str(ts_value).strip()
    try: return float(s)
    except:
        numbers = re.findall(r"\d+\.?\d*", s)
        return float(numbers[0]) if numbers else 0.0

def extract_frame_as_pil(video_path, seconds):
    cap = cv2.VideoCapture(video_path)
    cap.set(cv2.CAP_PROP_POS_MSEC, seconds * 1000)
    ret, frame = cap.read()
    cap.release()
    if ret:
        return PILImage.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
    return None

@st.cache_data
def generate_audio_bytes(text):
    try:
        if not text: return None
        tts = gTTS(text=text, lang='ja')
        fp = BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        return fp.read()
    except: return None

# Excel作成（画像合成機能付き）
def create_excel_file(steps, m_num, m_author, m_date, video_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "作業手順書"
    # (フォント設定等は省略せず記述)
    header_font = Font(bold=True, size=16)
    meta_font = Font(size=11)
    title_font = Font(bold=True, size=12)
    normal_font = Font(size=11)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws['A1'] = f"No: {m_num}"; ws['A1'].font = Font(bold=True, size=11)
    ws['C1'] = f"作成日: {m_date.strftime('%Y/%m/%d')}"; ws['C1'].alignment = Alignment(horizontal='right')
    ws['C2'] = f"作成者: {m_author}"; ws['C2'].alignment = Alignment(horizontal='right')
    ws.merge_cells('A3:C3'); ws['A3'] = "標準作業手順書"; ws['A3'].font = header_font; ws['A3'].alignment = Alignment(horizontal='center')

    start_row = 5
    for col, width, text in zip(['A','B','C'], [6, 45, 55], ["No.","作業画像","作業内容・手順"]):
        ws.column_dimensions[col].width = width
        cell = ws[f'{col}{start_row}']
        cell.value = text
        cell.font = title_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    current_row = start_row + 1
    for i, step in enumerate(steps, 1):
        ws.row_dimensions[current_row].height = 180
        ws[f'A{current_row}'].value = i
        ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'A{current_row}'].border = thin_border
        ws[f'B{current_row}'].border = thin_border
        
        # --- 画像合成 ---
        final_img = None
        if video_path:
            ts = clean_timestamp(step.get('timestamp', 0))
            if ts >= 0:
                final_img = extract_frame_as_pil(video_path, ts)

        if final_img and 'edited_image_data' in step and step['edited_image_data'] is not None:
            try:
                drawing_layer = PILImage.fromarray(step['edited_image_data'].astype('uint8'), 'RGBA')
                drawing_layer = drawing_layer.resize(final_img.size, PILImage.Resampling.LANCZOS)
                final_img.paste(drawing_layer, (0, 0), drawing_layer)
            except: pass

        if final_img:
            try:
                final_img.thumbnail((320, 240))
                img_byte_arr = BytesIO()
                final_img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                excel_img = ExcelImage(img_byte_arr)
                excel_img.anchor = f'B{current_row}'
                ws.add_image(excel_img)
            except: ws[f'B{current_row}'].value = "[画像エラー]"
        else: ws[f'B{current_row}'].value = "[画像なし]"

        cell_text = ws[f'C{current_row}']
        cell_text.value = f"【{step['title']}】\n\n{step['text']}"
        cell_text.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell_text.border = thin_border
        current_row += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. Gemini API ---
def process_video_with_gemini(video_path, api_key, selected_model):
    genai.configure(api_key=api_key)
    progress_bar = st.progress(0, text="準備中...")
    try:
        progress_bar.progress(10, text="📤 動画アップロード中...")
        video_file = genai.upload_file(path=video_path)
        while video_file.state.name == "PROCESSING":
            time.sleep(2); video_file = genai.get_file(video_file.name)
        if video_file.state.name == "FAILED": raise ValueError("動画処理失敗")

        progress_bar.progress(60, text=f"🤖 生成中 ({selected_model})...")
        model = genai.GenerativeModel(model_name=selected_model)
        prompt = """
        あなたは製造現場の熟練管理者です。動画を見て「標準作業手順書」を作成してください。
        JSON形式: [{"title": "...", "text": "...", "timestamp": 5.5},...]
        """
        response = model.generate_content(
            [video_file, prompt],
            generation_config={"response_mime_type": "application/json"}
        )
        progress_bar.progress(100, text="完了！"); time.sleep(1); progress_bar.empty()
        return json.loads(response.text)
    except Exception as e:
        st.error(f"エラー: {e}")
        return []

# --- 6. サーバー掃除 ---
def clear_api_storage(api_key):
    if not api_key: return
    try:
        genai.configure(api_key=api_key)
        files = list(genai.list_files())
        if not files: st.sidebar.success("ファイルなし"); return
        for f in files:
            try: genai.delete_file(f.name)
            except: pass
        st.sidebar.success(f"🧹 {len(files)}個削除完了")
    except: st.sidebar.error("削除失敗")

# --- 7. UI ---
with st.sidebar:
    st.header("🍌 Nano Banana")
    st.markdown("### Manufacturing AI Tools")
    st.divider()
    api_key = st.text_input("Google API Key", type="password")
    
    if api_key:
        available_models = get_available_models(api_key)
        # (モデル選択ロジックは簡略化して記述)
        final_model_name = st.selectbox("使用モデル", available_models, index=0)
        with st.expander("🛠️ メンテナンス"):
            if st.button("🗑️ ゴミ箱を空にする"): clear_api_storage(api_key)
    else: final_model_name = "gemini-1.5-flash"

    st.divider()
    st.header("📄 文書情報")
    manual_number = st.text_input("マニュアル番号", value="SOP-001")
    author_name = st.text_input("作成者", value="管理者")
    create_date = st.date_input("作成日", datetime.date.today())

st.title("📜 Nano Factory AI")
st.markdown("<p style='font-size: 1.3rem; font-weight: bold; color: #555;'>動画からマニュアルを自動生成・編集・Excel出力まで一気通貫で行います。</p>", unsafe_allow_html=True)

if "edit_mode" not in st.session_state: st.session_state.edit_mode = "list" # list or draw
if "manual_steps" not in st.session_state: st.session_state.manual_steps = None

uploaded_file = st.file_uploader("動画アップロード", type=["mp4", "mov"], label_visibility="collapsed")

if uploaded_file:
    temp_filename = "temp_video.mp4"
    if not os.path.exists(temp_filename): # 再アップロード回避
        with open(temp_filename, "wb") as f:
            while True:
                chunk = uploaded_file.read(1024*1024)
                if not chunk: break
                f.write(chunk)

    if st.session_state.edit_mode == "list":
        # === モード1：リスト表示 & 秒数調整 ===
        st.subheader("🎥 現場動画（元データ）")
        st.video(temp_filename)
        st.divider()

        if st.button("AI解析を実行する", type="primary"):
            if not api_key: st.error("APIキーが必要です")
            else:
                with st.spinner("AI解析中..."):
                    steps = process_video_with_gemini(temp_filename, api_key, final_model_name)
                    if steps:
                        st.session_state.manual_steps = steps
                        st.rerun()

        if st.session_state.manual_steps:
            st.subheader("📝 編集 & プレビュー")
            st.info("秒数を調整して、ベストな画像を選んでください。お絵かきは「次へ」ボタンを押してから行います。")
            
            steps = st.session_state.manual_steps
            for i, step in enumerate(steps):
                with st.container():
                    st.markdown(f"#### 手順 {i+1}")
                    c1, c2 = st.columns([1.5, 1])
                    with c1:
                        ts = clean_timestamp(step.get('timestamp', 0))
                        new_ts = st.number_input(f"秒数 #{i+1}", value=ts, step=0.1, format="%.1f", key=f"ts_{i}")
                        img = extract_frame_as_pil(temp_filename, new_ts)
                        if img: st.image(img, use_container_width=True)
                        steps[i]['timestamp'] = new_ts
                    with c2:
                        steps[i]['title'] = st.text_input(f"見出し #{i+1}", step['title'], key=f"ti_{i}")
                        steps[i]['text'] = st.text_area(f"説明 #{i+1}", step['text'], height=150, key=f"tx_{i}")
                    st.divider()
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                # 確定して次のモードへ
                if st.button("🎨 画像を編集（お絵かき）する", type="primary", use_container_width=True):
                    st.session_state.edit_mode = "draw"
                    st.rerun()
            with col_btn2:
                # そのまま出力
                excel_data = create_excel_file(steps, manual_number, author_name, create_date, temp_filename)
                st.download_button("📥 そのままExcel出力", excel_data, f"{manual_number}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    elif st.session_state.edit_mode == "draw":
        # === モード2：お絵かき集中モード ===
        st.subheader("🎨 画像編集モード")
        st.info("1枚ずつ選択して、矢印や枠線を描き込んでください。")
        
        steps = st.session_state.manual_steps
        
        # 編集する手順を選ぶセレクトボックス
        step_options = [f"手順 {i+1}: {s['title']}" for i, s in enumerate(steps)]
        selected_option = st.selectbox("編集する画像を選択:", step_options)
        selected_index = step_options.index(selected_option)
        
        # ツールバー
        t1, t2, t3 = st.columns([1,1,2])
        with t1: mode = st.selectbox("ツール", ["rect", "circle", "line", "text", "transform"], key="draw_tool")
        with t2: color = st.color_picker("色", "#FF0000", key="draw_color")
        with t3: width = st.slider("太さ", 1, 10, 3, key="draw_width")

        # キャンバス表示（1つだけ表示するので軽快！）
        target_step = steps[selected_index]
        ts = clean_timestamp(target_step.get('timestamp', 0))
        bg_img = extract_frame_as_pil(temp_filename, ts)
        
        if bg_img:
            # 高画質すぎると重いのでリサイズして表示（保存時は合成で綺麗にする）
            display_img = bg_img.copy()
            display_img.thumbnail((800, 800))
            
            # 既存の編集データがあれば読み込む
            initial_data = target_step.get('edited_image_data')
            
            canvas_result = st_canvas(
                fill_color="rgba(255, 165, 0, 0.1)",
                stroke_width=width, stroke_color=color,
                background_image=display_img,
                update_streamlit=True,
                height=400, # 大きく表示
                drawing_mode=mode,
                initial_drawing=None, # 再編集は難しいので簡易実装
                key=f"canvas_editor_{selected_index}", # キーを変えてリセット防止
                display_toolbar=True,
            )
            
            # 描画データを保存
            if canvas_result.image_data is not None:
                steps[selected_index]['edited_image_data'] = canvas_result.image_data
                st.success("✅ 編集内容を一時保存中...")

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("↩️ リストに戻る"):
                st.session_state.edit_mode = "list"
                st.rerun()
        with c2:
            excel_data = create_excel_file(steps, manual_number, author_name, create_date, temp_filename)
            st.download_button("📥 編集完了！Excelをダウンロード", excel_data, f"{manual_number}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
