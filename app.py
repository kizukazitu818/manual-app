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
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
from gtts import gTTS

# --- 1. アプリ設定 & デザイン定義 ---
st.set_page_config(
    page_title="SOP Generator Enterprise",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ★視認性改善CSS：どんな環境でも「白背景・黒文字」を強制する
st.markdown("""
    <style>
    /* ベースの強制上書き */
    html, body, [class*="css"] {
        font-family: 'Inter', 'Helvetica Neue', Arial, sans-serif;
    }
    
    /* アプリ全体の背景と文字色 */
    .stApp {
        background-color: #f1f5f9 !important;
        color: #1e293b !important;
    }

    /* サイドバー */
    [data-testid="stSidebar"] {
        background-color: #1e293b !important;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span, [data-testid="stSidebar"] p {
        color: #f8fafc !important; /* サイドバー内の文字は白 */
    }
    
    /* 見出し（黒） */
    h1, h2, h3, h4, h5, h6 {
        color: #0f172a !important;
        font-weight: 700 !important;
    }
    
    /* 本文テキスト（黒） */
    p, div, span, label, li {
        color: #334155 !important;
    }
    
    /* カードデザイン */
    .step-card {
        background-color: #ffffff !important;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
        border: 1px solid #cbd5e1;
    }

    /* 入力フォームの文字色強制（重要！） */
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-baseweb="select"] div {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border-color: #cbd5e1 !important;
    }
    
    /* アップローダー等の文字色 */
    [data-testid="stFileUploader"] label {
        color: #0f172a !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
    }

    /* ボタン */
    div.stButton > button:first-child {
        background-color: #2563eb !important;
        color: #ffffff !important; /* ボタン文字は白 */
        border: none;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2);
    }
    div.stButton > button:first-child:hover {
        background-color: #1d4ed8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ヘッダーエリア
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("🏭 SOP Generator Enterprise")
    st.markdown("**映像解析AIによる、標準作業手順書（SOP）自動生成プラットフォーム**")
with col_h2:
    st.markdown("""
        <div style='background-color:white; padding:10px; border-radius:8px; text-align:center; border:1px solid #ddd;'>
            <small style='color:#64748b !important; font-weight:bold;'>SYSTEM STATUS</small><br>
            <span style='color:#10b981 !important; font-weight:bold;'>● ONLINE</span>
        </div>
    """, unsafe_allow_html=True)

# --- 2. サイドバー設定 ---
with st.sidebar:
    st.markdown("### ⚙️ 設定パネル")
    api_key = st.text_input("Google API Key", type="password")
    
    st.divider()
    
    st.markdown("### 🧠 AIエンジンの選択")
    model_options = [
        "gemini-2.0-flash-exp", 
        "gemini-1.5-pro",       
        "gemini-1.5-flash",     
        "gemini-1.0-pro"        
    ]
    selected_model = st.selectbox("Model", model_options, index=0)

    st.divider()
    
    st.markdown("### 📄 ドキュメント情報")
    manual_number = st.text_input("文書番号", value="SOP-2026-001")
    author_name = st.text_input("作成者", value="管理者")
    create_date = st.date_input("作成日", datetime.date.today())

# --- 3. データ処理用ヘルパー関数 ---
def clean_timestamp(ts_value):
    if ts_value is None: return 0.0
    if isinstance(ts_value, (int, float)): return float(ts_value)
    s = str(ts_value).strip()
    try: return float(s)
    except ValueError:
        if ":" in s:
            parts = s.split(":")
            if len(parts) == 2:
                try: return float(parts[0]) * 60 + float(parts[1])
                except: pass
        numbers = re.findall(r"\d+\.?\d*", s)
        if numbers: return float(numbers[0])
    return 0.0

def extract_frame_for_web(video_path, seconds):
    cap = cv2.VideoCapture(video_path)
    cap.set(cv2.CAP_PROP_POS_MSEC, seconds * 1000)
    ret, frame = cap.read()
    cap.release()
    if ret: return cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    return None

def extract_frame_for_excel(video_path, seconds):
    frame_rgb = extract_frame_for_web(video_path, seconds)
    if frame_rgb is not None: return PILImage.fromarray(frame_rgb)
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
    except Exception as e: return None

# --- 4. Excel作成関数 ---
def create_excel_file(steps, m_num, m_author, m_date, video_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "作業手順書"

    header_font = Font(bold=True, size=16, name='Meiryo UI')
    title_font = Font(bold=True, size=12, name='Meiryo UI')
    normal_font = Font(size=11, name='Meiryo UI')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    alignment_center = Alignment(horizontal='center', vertical='center')
    alignment_left = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws['A1'] = f"No: {m_num}"
    ws['C1'] = f"作成日: {m_date.strftime('%Y/%m/%d')}"
    ws['C1'].alignment = Alignment(horizontal='right')
    ws['C2'] = f"作成者: {m_author}"
    ws['C2'].alignment = Alignment(horizontal='right')
    ws.merge_cells('A3:C3')
    ws['A3'] = "標準作業手順書"
    ws['A3'].font = header_font
    ws['A3'].alignment = alignment_center

    start_row = 5
    headers = ["No.", "作業画像", "作業内容・手順"]
    widths = [6, 45, 55]
    for i, (h, w) in enumerate(zip(headers, widths)):
        col = chr(65 + i)
        ws[f'{col}{start_row}'] = h
        ws.column_dimensions[col].width = w
        ws[f'{col}{start_row}'].font = title_font
        ws[f'{col}{start_row}'].border = thin_border
        ws[f'{col}{start_row}'].alignment = alignment_center

    current_row = start_row + 1
    for i, step in enumerate(steps, 1):
        ws.row_dimensions[current_row].height = 180
        
        cell = ws[f'A{current_row}']
        cell.value = i
        cell.alignment = alignment_center
        cell.border = thin_border
        
        cell_img = ws[f'B{current_row}']
        cell_img.border = thin_border
        ts = clean_timestamp(step.get('timestamp', 0))
        if video_path and ts >= 0:
            try:
                pil_img = extract_frame_for_excel(video_path, ts)
                if pil_img:
                    pil_img.thumbnail((320, 240))
                    img_byte_arr = BytesIO()
                    pil_img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    excel_img = ExcelImage(img_byte_arr)
                    excel_img.anchor = f'B{current_row}'
                    ws.add_image(excel_img)
                else: cell_img.value = "[画像なし]"
            except: cell_img.value = "[エラー]"
        
        cell_text = ws[f'C{current_row}']
        cell_text.value = f"【{step['title']}】\n\n{step['text']}"
        cell_text.alignment = alignment_left
        cell_text.border = thin_border
        cell_text.font = normal_font
        
        current_row += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. Gemini API処理 ---
def process_video_with_gemini(video_path, api_key, model_name):
    genai.configure(api_key=api_key)
    status_text = st.status("🚀 AIエージェントを起動中...", expanded=True)
    try:
        status_text.write("📤 映像データをクラウドへ転送しています...")
        video_file = genai.upload_file(path=video_path)
        
        while video_file.state.name == "PROCESSING":
            status_text.write("⏳ 映像をフレーム単位で解析中...")
            time.sleep(2)
            video_file = genai.get_file(video_file.name)
            
        if video_file.state.name == "FAILED": raise ValueError("動画処理失敗")

        status_text.write(f"🧠 {model_name} が作業手順を構造化しています...")
        model = genai.GenerativeModel(model_name=model_name)
        
        prompt = """
        あなたは製造現場の熟練管理者です。添付の動画を見て、新人作業員のための「標準作業手順書」を作成してください。
        以下のJSON形式で出力してください:
        [{"title": "見出し", "text": "詳細手順", "timestamp": 5.5}, ...]
        注意: timestampは必ず秒数(数値)のみ。
        """
        response = model.generate_content([video_file, prompt], generation_config={"response_mime_type": "application/json"})
        
        status_text.update(label="✅ 生成完了！", state="complete", expanded=False)
        return json.loads(response.text)
    except Exception as e:
        status_text.update(label="❌ エラー発生", state="error")
        st.error(f"Error: {e}")
        return []

# --- 6. メインエリア ---

# ファイルアップロードエリア
st.markdown("### 1. 映像データの入力")
with st.container():
    st.markdown("""
        <div style='background-color:white; padding:20px; border-radius:10px; border: 2px dashed #cbd5e1; text-align:center;'>
            <p style='margin:0; color:#64748b !important;'>↓ ここに作業動画ファイルをドロップしてください</p>
        </div>
    """, unsafe_allow_html=True)
    # ★修正ポイント：ラベルを追加して警告を消去、label_visibilityで隠す
    uploaded_file = st.file_uploader("動画アップロード", type=["mp4", "mov"], label_visibility="collapsed")

if uploaded_file is not None:
    temp_filename = "temp_video.mp4"
    # メモリ節約読み込み
    with open(temp_filename, "wb") as f:
        while True:
            chunk = uploaded_file.read(1024 * 1024)
            if not chunk: break
            f.write(chunk)

    col_v1, col_v2 = st.columns([2, 1])
    
    with col_v1:
        st.video(temp_filename)
        
    with col_v2:
        st.markdown("### アクション")
        st.info("動画の内容をAIが解析し、手順書の下書きを作成します。")
        
        if "manual_steps" not in st.session_state:
            st.session_state.manual_steps = None

        if st.button("✨ 自動解析を実行する", type="primary", use_container_width=True):
            if not api_key:
                st.error("APIキーを入力してください")
            else:
                steps = process_video_with_gemini(temp_filename, api_key, selected_model)
                st.session_state.manual_steps = steps
                st.rerun()

    # --- 編集エリア ---
    if st.session_state.manual_steps:
        st.divider()
        st.markdown("### 2. 手順の編集・構成")
        
        steps = st.session_state.manual_steps
        
        with st.form("edit_form"):
            for i, step in enumerate(steps):
                st.markdown(f"""
                <div class="step-card">
                    <h4 style="margin-top:0; color:#3b82f6 !important;">STEP {i+1}</h4>
                </div>
                """, unsafe_allow_html=True)
                
                c1, c2 = st.columns([1, 2])
                
                with c1:
                    current_ts = clean_timestamp(step.get('timestamp', 0.0))
                    new_ts = st.number_input(f"⏱ 秒数 (Step {i+1})", min_value=0.0, value=current_ts, step=0.1, key=f"ts_{i}")
                    
                    frame = extract_frame_for_web(temp_filename, new_ts)
                    if frame is not None:
                        st.image(frame, use_container_width=True, caption=f"{new_ts}秒地点")
                    
                    steps[i]['timestamp'] = new_ts
                    
                with c2:
                    steps[i]['title'] = st.text_input(f"見出し (Step {i+1})", value=step['title'], key=f"t_{i}")
                    steps[i]['text'] = st.text_area(f"説明文 (Step {i+1})", value=step['text'], height=150, key=f"d_{i}")

            st.markdown("<br>", unsafe_allow_html=True)
            submitted = st.form_submit_button("✅ 編集を確定してプレビューへ進む", use_container_width=True)
            
            if submitted:
                st.success("内容を更新しました！")

        # --- 最終プレビュー ---
        st.divider()
        st.markdown("### 3. 出力プレビュー")
        
        with st.container():
            st.markdown(f"""
            <div style="background-color:white; padding:40px; border:1px solid #ddd; border-radius:4px;">
                <h2 style="text-align:center; border-bottom:2px solid #333; padding-bottom:10px;">標準作業手順書</h2>
                <div style="display:flex; justify-content:space-between; color:#666; margin-bottom:20px;">
                    <span style='color:#333 !important;'>No: {manual_number}</span>
                    <span style='color:#333 !important;'>作成: {author_name} ({create_date})</span>
                </div>
            """, unsafe_allow_html=True)
            
            for i, step in enumerate(steps, 1):
                c_p1, c_p2, c_p3 = st.columns([0.2, 1, 2])
                with c_p1: st.markdown(f"**{i}**")
                with c_p2:
                    ts = clean_timestamp(step.get('timestamp', 0))
                    f = extract_frame_for_web(temp_filename, ts)
                    if f is not None: st.image(f, use_container_width=True)
                with c_p3:
                    st.markdown(f"**{step['title']}**")
                    st.markdown(step['text'])
                    
                    txt = f"手順{i}。{step['title']}。{step['text']}"
                    aud = generate_audio_bytes(txt)
                    if aud: st.audio(aud, format='audio/mp3')
                
                st.divider()
                
            st.markdown("</div>", unsafe_allow_html=True)

        excel_data = create_excel_file(steps, manual_number, author_name, create_date, temp_filename)
        st.download_button(
            label="📥 Excelファイルとしてダウンロード",
            data=excel_data,
            file_name=f"{manual_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
