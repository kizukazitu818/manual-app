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

# --- 1. アプリ全体の基本設定 & デザイン（視認性重視） ---
st.set_page_config(
    page_title="Auto-Manual Producer Pro",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ★視認性を高めるカスタムCSS（文字を黒く、背景を優しく）
st.markdown("""
    <style>
    /* 全体の背景を薄いグレーに */
    .stApp {
        background-color: #f4f6f9;
        color: #333333;
    }
    
    /* サイドバーの背景 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
    }
    
    /* 入力フォームやカードのスタイル（白背景に黒文字） */
    .stForm, div[data-testid="stExpander"] {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border: 1px solid #e0e0e0;
    }
    
    /* 文字色の強制指定（白飛び防止） */
    h1, h2, h3, h4, h5, h6, p, label, span, div {
        color: #1f2937 !important; 
    }
    
    /* プレビュー内のテキストは黒く */
    [data-testid="stMarkdownContainer"] p {
        color: #333333 !important;
    }

    /* ボタンのカスタマイズ */
    div.stButton > button:first-child {
        background-color: #2563eb;
        color: white !important; /* ボタンの文字だけは白 */
        font-weight: bold;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 1rem;
    }
    div.stButton > button:first-child:hover {
        background-color: #1d4ed8;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🛠️ Auto-Manual Producer Pro")
st.markdown("##### 現場動画から、プロ品質の標準作業手順書（SOP）を瞬時に生成。")

# --- 2. サイドバー設定 ---
with st.sidebar:
    st.header("⚙️ システム設定")
    
    # APIキー入力
    api_key = st.text_input("Google API Key", type="password", help="Geminiを利用するためのキーを入力")
    
    st.divider()
    
    # モデル選択
    st.subheader("🧠 AIモデル選択")
    model_options = [
        "gemini-2.0-flash-exp", 
        "gemini-1.5-pro",       
        "gemini-1.5-flash",     
        "gemini-1.0-pro"        
    ]
    selected_model = st.selectbox(
        "使用するモデル", 
        model_options,
        index=0
    )

    st.divider()
    
    st.header("📄 文書プロパティ")
    manual_number = st.text_input("文書番号 (No)", value="SOP-001")
    author_name = st.text_input("作成者", value="管理者")
    create_date = st.date_input("作成日", datetime.date.today())

# --- 3. データ処理用ヘルパー関数群 ---
def clean_timestamp(ts_value):
    """
    AIが '0:31' や 'approx 5s' などの形式で返してきた場合に
    強制的に秒数(float)に変換するフィルター関数
    """
    if ts_value is None: return 0.0
    if isinstance(ts_value, (int, float)): return float(ts_value)
    
    s = str(ts_value).strip()
    try:
        return float(s)
    except ValueError:
        if ":" in s:
            parts = s.split(":")
            if len(parts) == 2:
                try:
                    return float(parts[0]) * 60 + float(parts[1])
                except: pass
        numbers = re.findall(r"\d+\.?\d*", s)
        if numbers:
            return float(numbers[0])
    return 0.0

def extract_frame_for_web(video_path, seconds):
    """Web表示用に高速にフレームを切り出す"""
    cap = cv2.VideoCapture(video_path)
    cap.set(cv2.CAP_PROP_POS_MSEC, seconds * 1000)
    ret, frame = cap.read()
    cap.release()
    if ret:
        return cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    return None

def extract_frame_for_excel(video_path, seconds):
    """Excel貼り付け用にフレームを切り出す"""
    frame_rgb = extract_frame_for_web(video_path, seconds)
    if frame_rgb is not None:
        return PILImage.fromarray(frame_rgb)
    return None

@st.cache_data
def generate_audio_bytes(text):
    """テキストから音声を生成してバイナリデータで返す"""
    try:
        if not text: return None
        tts = gTTS(text=text, lang='ja')
        fp = BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        return fp.read()
    except Exception as e:
        return None

# --- 4. Excel作成関数 ---
def create_excel_file(steps, m_num, m_author, m_date, video_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "作業手順書"

    header_font = Font(bold=True, size=16)
    meta_font = Font(size=11)
    title_font = Font(bold=True, size=12)
    normal_font = Font(size=11)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    ws['A1'] = f"No: {m_num}"
    ws['A1'].font = Font(bold=True, size=11)
    ws['C1'] = f"作成日: {m_date.strftime('%Y/%m/%d')}"
    ws['C1'].font = meta_font
    ws['C1'].alignment = Alignment(horizontal='right')
    ws['C2'] = f"作成者: {m_author}"
    ws['C2'].font = meta_font
    ws['C2'].alignment = Alignment(horizontal='right')
    ws.merge_cells('A3:C3')
    ws['A3'] = "標準作業手順書"
    ws['A3'].font = header_font
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    start_row = 5
    ws[f'A{start_row}'] = "No."
    ws[f'B{start_row}'] = "作業画像"
    ws[f'C{start_row}'] = "作業内容・手順"
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 55
    for col in ['A', 'B', 'C']:
        cell = ws[f'{col}{start_row}']
        cell.font = title_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    current_row = start_row + 1
    for i, step in enumerate(steps, 1):
        ws.row_dimensions[current_row].height = 180
        cell_no = ws[f'A{current_row}']
        cell_no.value = i
        cell_no.alignment = Alignment(horizontal='center', vertical='center')
        cell_no.border = thin_border
        
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
                else:
                    cell_img.value = "[画像取得失敗]"
            except Exception as e:
                cell_img.value = f"[エラー]"
        else:
            cell_img.value = "[画像なし]"

        cell_text = ws[f'C{current_row}']
        cell_text.value = f"【{step['title']}】\n\n{step['text']}"
        cell_text.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell_text.border = thin_border
        cell_text.font = normal_font
        current_row += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. Gemini API処理 ---
def process_video_with_gemini(video_path, api_key, model_name):
    genai.configure(api_key=api_key)
    status_text = st.empty()
    try:
        status_text.info(f"📤 動画をAIサーバーにアップロード中... (モデル: {model_name})")
        video_file = genai.upload_file(path=video_path)
        while video_file.state.name == "PROCESSING":
            status_text.info("⏳ AIが動画を処理しています...")
            time.sleep(2)
            video_file = genai.get_file(video_file.name)
        if video_file.state.name == "FAILED": raise ValueError("動画処理失敗")

        status_text.info(f"🤖 マニュアルを生成中... ({model_name})")
        
        model = genai.GenerativeModel(model_name=model_name)
        
        prompt = """
        あなたは製造現場の熟練管理者です。添付の動画を見て、新人作業員のための「標準作業手順書」を作成してください。
        以下のJSON形式で出力してください:
        [
            {"title": "手順の見出し", "text": "具体的な作業内容。", "timestamp": 5.5},...
        ]
        注意点: 
        - timestampは必ず「秒数（数値）」だけにしてください。（例: 5.5）
        - 専門用語を正しく使い、曖昧な指示は具体化すること。
        """
        response = model.generate_content([video_file, prompt], generation_config={"response_mime_type": "application/json"})
        status_text.success("完了！下の編集エリアで内容を確認してください。")
        return json.loads(response.text)
    except Exception as e:
        st.error(f"エラー: {e}")
        return []

# --- 6. メインエリア ---
uploaded_file = st.file_uploader("📂 作業動画をここにドラッグ＆ドロップ", type=["mp4", "mov"], help="AIが解析する動画ファイルをアップロードしてください")

if uploaded_file is not None:
    temp_filename = "temp_video.mp4"
    
    # 【★重要修正】メモリを節約するために、少しずつファイルを保存する方式に変更
    # これにより "Connection Reset" (OOM) エラーを防ぎます
    with open(temp_filename, "wb") as f:
        while True:
            chunk = uploaded_file.read(1024 * 1024) # 1MBずつ読み込む
            if not chunk:
                break
            f.write(chunk)

    with st.expander("⚙️ プレビュー表示設定"):
        col_size1, col_size2 = st.columns(2)
        with col_size1:
            video_width = st.slider("動画プレイヤー幅 (%)", 10, 100, 60)
        with col_size2:
            img_width = st.slider("編集画像幅 (%)", 10, 100, 100)

    st.markdown("### 🎥 現場動画プレビュー")
    left, center, right = st.columns([1, 2, 1])
    if video_width > 50:
        left, center, right = st.columns([0.1, 1, 0.1])
        
    with center:
        st.video(temp_filename)
    
    st.divider()
    
    st.markdown("### 📝 手順作成・編集")
    
    if "manual_steps" not in st.session_state:
        st.session_state.manual_steps = None

    if st.button("🚀 AI解析を開始する", type="primary", use_container_width=True):
        if not api_key:
            st.error("⚠️ 左側の設定メニューでAPIキーを入力してください！")
        else:
            with st.spinner(f"AI ({selected_model}) が動画を解析中..."):
                steps = process_video_with_gemini(temp_filename, api_key, selected_model)
                st.session_state.manual_steps = steps
                st.rerun()
    
    # --- 編集エリア ---
    if st.session_state.manual_steps:
        steps = st.session_state.manual_steps
        
        st.info("💡 以下のフォームで内容を微調整できます。画像位置（秒数）を変えると、リアルタイムに画像が切り替わります。")

        with st.form("edit_form"):
            for i, step in enumerate(steps):
                st.markdown(f"#### Step {i+1}")
                col_ratio_img = 1 + (img_width / 100)
                col_ratio_text = 4 - (img_width / 100)
                col_img, col_text = st.columns([col_ratio_img, col_ratio_text])
                
                with col_img:
                    current_ts = clean_timestamp(step.get('timestamp', 0.0))
                    new_timestamp = st.number_input(
                        f"📷 画像位置(秒)", min_value=0.0, value=current_ts, step=0.1, format="%.1f", key=f"ts_{i}"
                    )
                    frame_rgb = extract_frame_for_web(temp_filename, new_timestamp)
                    if frame_rgb is not None:
                        st.image(frame_rgb, caption=f"{new_timestamp}秒時点", use_container_width=True)
                    steps[i]['timestamp'] = new_timestamp

                with col_text:
                    new_title = st.text_input(f"見出し", value=step['title'], key=f"title_{i}")
                    new_text = st.text_area(f"詳細手順", value=step['text'], key=f"text_{i}", height=120)
                    steps[i]['title'] = new_title
                    steps[i]['text'] = new_text
                st.divider()
            
            submitted = st.form_submit_button("✅ 編集を確定してプレビュー", use_container_width=True)
            if submitted:
                st.success("編集内容を更新しました！下にスクロールして完成形を確認してください。")

        # --- プレビュー ---
        st.markdown("### 📑 完成プレビュー & 音声確認")
        with st.container(border=True): 
            col_ph1, col_ph2 = st.columns([1,1])
            with col_ph1:
                st.markdown(f"**No:** {manual_number}")
            with col_ph2:
                st.markdown(f"<div style='text-align: right'>作成日: {create_date}<br>作成者: {author_name}</div>", unsafe_allow_html=True)
            
            st.markdown("<h2 style='text-align: center; border-bottom: 2px solid #ddd;'>標準作業手順書</h2>", unsafe_allow_html=True)
            st.write("") 
            
            for i, step in enumerate(steps, 1):
                p_col1, p_col2, p_col3 = st.columns([0.3, 3, 4])
                with p_col1: st.markdown(f"<h3 style='color: #888;'>{i}</h3>", unsafe_allow_html=True)
                with p_col2:
                    ts = clean_timestamp(step.get('timestamp', 0))
                    if temp_filename:
                        frame_rgb = extract_frame_for_web(temp_filename, ts)
                        if frame_rgb is not None:
                            st.image(frame_rgb, use_container_width=True, output_format="JPEG")
                with p_col3:
                    st.markdown(f"#### {step['title']}")
                    st.write(step['text'])
                    
                    read_text = f"手順{i}。{step['title']}。{step['text']}"
                    audio_bytes = generate_audio_bytes(read_text)
                    if audio_bytes:
                        st.audio(audio_bytes, format='audio/mp3')

                st.markdown("---")

        excel_data = create_excel_file(steps, manual_number, author_name, create_date, temp_filename)
        st.download_button(
            label="📥 Excelファイルを出力する",
            data=excel_data,
            file_name=f"{manual_number}_manual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
