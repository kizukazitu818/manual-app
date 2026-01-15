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

# --- 1. アプリ全体の基本設定 ---
st.set_page_config(
    page_title="Auto-Manual Producer",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ★ここが新機能！UIを強制的にカスタマイズするCSS★
# アップロード欄の背景を「淡い朱色」にし、枠線を「濃い朱色」にします
st.markdown("""
    <style>
    /* ファイルアップロード欄の背景色 */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFF0F0; /* 淡い朱色 */
        border: 1px dashed #FF4B4B; /* 枠線を朱色に */
    }
    /* サイドバーの背景色（念のためCSSでも指定） */
    [data-testid="stSidebar"] {
        background-color: #FFF0F0;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🛠️ Auto-Manual Producer (AMP)")
st.caption("動画からマニュアルを自動生成・編集・Excel出力まで一気通貫で行います。")

# --- 2. モデルリスト取得関数 ---
@st.cache_data(ttl=600)
def get_available_models(api_key):
    default_models = ["gemini-1.5-flash"]
    if not api_key: return default_models
    try:
        genai.configure(api_key=api_key)
        models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                models.append(name)
        models.sort()
        return models if models else default_models
    except Exception:
        return default_models

# --- 3. サイドバー設定 ---
with st.sidebar:
    st.header("設定")
    api_key = st.text_input("Google API Key", type="password")
    
    st.divider()
    
    st.header("🧠 AIモデル選択")
    
    if api_key:
        available_models = get_available_models(api_key)
        
        st.subheader("① 作成目的を選ぶ")
        scenario = st.radio(
            "どのような視点の手順書を作成しますか？",
            [
                "🔧 メカニック視点（点検・保全用）",
                "🛡️ 安全管理者視点（教育・ルール用）",
                "📹 解析・記録視点（動画リンク用）",
                "🚀 標準（バランス型）"
            ],
            index=3,
            help="選んだ視点に合わせて、最適なAIモデルが自動的に推奨されます。"
        )

        recommended_keyword = ""
        if "mechanic" in scenario or "メカニック" in scenario:
            recommended_keyword = "gemini-2.5"
            st.info("💡 Point: 部品の劣化や緩みなど、設備の状態を細かく描写します。")
        elif "safety" in scenario or "安全管理" in scenario:
            recommended_keyword = "gemini-3"
            st.info("💡 Point: 指差し確認や安全タグなど、ルールや安全行動を重視します。")
        elif "robotics" in scenario or "解析・記録" in scenario:
            recommended_keyword = "robotics"
            st.info("💡 Point: 「(00:15-00:20)」のように正確なタイムスタンプを記録します。")
        else:
            recommended_keyword = "gemini-1.5-flash"

        default_index = 0
        for i, model_name in enumerate(available_models):
            if recommended_keyword in model_name:
                default_index = i
                break
        
        st.subheader("② 使用するモデルを確認")
        final_model_name = st.selectbox(
            "実際に使用するモデル（自動選択されます）",
            available_models,
            index=default_index
        )

    else:
        st.info("APIキーを入力すると、モデル選択メニューが表示されます。")
        final_model_name = "gemini-1.5-flash"

    st.divider()
    st.header("📄 文書情報")
    manual_number = st.text_input("マニュアル番号", value="SOP-001")
    author_name = st.text_input("作成者", value="管理者")
    create_date = st.date_input("作成日", datetime.date.today())

# --- 4. データ処理用ヘルパー関数群 ---
def clean_timestamp(ts_value):
    if ts_value is None: return 0.0
    if isinstance(ts_value, (int, float)): return float(ts_value)
    s = str(ts_value).strip()
    try:
        return float(s)
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
    if ret:
        return cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    return None

def extract_frame_for_excel(video_path, seconds):
    frame_rgb = extract_frame_for_web(video_path, seconds)
    if frame_rgb is not None:
        return PILImage.fromarray(frame_rgb)
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
    except Exception:
        return None

# --- 5. Excel作成関数 ---
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
            except Exception:
                cell_img.value = "[画像エラー]"
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

# --- 6. Gemini API処理 ---
def process_video_with_gemini(video_path, api_key, selected_model):
    genai.configure(api_key=api_key)
    
    progress_bar = st.progress(0, text="準備中...")
    
    try:
        progress_bar.progress(10, text="📤 動画をAIサーバーにアップロード中...")
        video_file = genai.upload_file(path=video_path)
        
        while video_file.state.name == "PROCESSING":
            progress_bar.progress(30, text="⏳ AI側で動画を処理しています...（数秒〜数分）")
            time.sleep(2)
            video_file = genai.get_file(video_file.name)
            
        if video_file.state.name == "FAILED":
            raise ValueError("動画の処理に失敗しました。")

        progress_bar.progress(60, text=f"🤖 マニュアルを生成中...（モデル: {selected_model}）")
        
        model = genai.GenerativeModel(model_name=selected_model)
        
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
        response = model.generate_content(
            [video_file, prompt],
            generation_config={"response_mime_type": "application/json"}
        )
        
        progress_bar.progress(100, text="完了！")
        time.sleep(1)
        progress_bar.empty()
        
        return json.loads(response.text)

    except Exception as e:
        st.error(f"エラー: {e}")
        return []

# --- 7. メインエリア ---
uploaded_file = st.file_uploader("作業動画をアップロードしてください", type=["mp4", "mov"])

if uploaded_file is not None:
    temp_filename = "temp_video.mp4"
    with open(temp_filename, "wb") as f: f.write(uploaded_file.read())

    with st.expander("⚙️ 表示サイズ調整"):
        col_size1, col_size2 = st.columns(2)
        with col_size1:
            video_width = st.slider("動画プレイヤーのサイズ (%)", 10, 100, 50)
        with col_size2:
            img_width = st.slider("編集画像のサイズ (%)", 10, 100, 100)

    st.subheader("🎥 現場動画（元データ）")
    
    left_padding = (100 - video_width) / 2
    right_padding = (100 - video_width) / 2
    cols = st.columns([max(0.1, left_padding), video_width, max(0.1, right_padding)])
    with cols[1]:
        st.video(uploaded_file)
    
    st.divider()
    
    st.subheader("📝 編集 & プレビュー")
    
    if "manual_steps" not in st.session_state:
        st.session_state.manual_steps = None

    if st.button("AI解析を実行する", type="primary"):
        if not api_key:
            st.error("⚠️ APIキーを入力してください！")
        else:
            with st.spinner(f"AIエージェントを起動中（モデル: {final_model_name}）..."):
                steps = process_video_with_gemini(temp_filename, api_key, final_model_name)
                st.session_state.manual_steps = steps
                st.rerun()
    
    # --- 編集エリア ---
    if st.session_state.manual_steps:
        steps = st.session_state.manual_steps
        
        st.markdown(f"### ✍️ 手順の編集（使用モデル: {final_model_name}）")
        with st.form("edit_form"):
            for i, step in enumerate(steps):
                st.markdown(f"#### 手順 {i+1}")
                col_ratio_img = 1 + (img_width / 100)
                col_ratio_text = 4 - (img_width / 100)
                col_img, col_text = st.columns([col_ratio_img, col_ratio_text])
                
                with col_img:
                    current_ts = clean_timestamp(step.get('timestamp', 0.0))
                    new_timestamp = st.number_input(
                        f"画像位置(秒)", min_value=0.0, value=current_ts, step=0.1, format="%.1f", key=f"ts_{i}"
                    )
                    frame_rgb = extract_frame_for_web(temp_filename, new_timestamp)
                    if frame_rgb is not None:
                         st.image(frame_rgb, caption=f"{new_timestamp}秒時点", width=None, use_container_width=True)
                    steps[i]['timestamp'] = new_timestamp

                with col_text:
                    new_title = st.text_input(f"見出し", value=step['title'], key=f"title_{i}")
                    new_text = st.text_area(f"説明", value=step['text'], key=f"text_{i}", height=150)
                    steps[i]['title'] = new_title
                    steps[i]['text'] = new_text
                st.divider()
            
            submitted = st.form_submit_button("✅ 編集内容を確定してプレビューへ")
            if submitted:
                st.success("内容を更新しました！下のプレビューを確認してください。")

        # --- プレビュー ---
        st.markdown("### 📄 完成イメージ（プレビュー & 音声確認）")
        with st.container(border=True): 
            st.markdown(f"**No:** {manual_number}　　**作成日:** {create_date}　　**作成者:** {author_name}")
            st.markdown("## 標準作業手順書")
            st.divider()
            
            for i, step in enumerate(steps, 1):
                p_col1, p_col2, p_col3 = st.columns([0.5, 3, 4])
                with p_col1: st.markdown(f"### {i}")
                with p_col2:
                    ts = clean_timestamp(step.get('timestamp', 0))
                    if temp_filename:
                        frame_rgb = extract_frame_for_web(temp_filename, ts)
                        if frame_rgb is not None:
                            st.image(frame_rgb, use_container_width=True)
                with p_col3:
                    st.markdown(f"#### {step['title']}")
                    st.write(step['text'])
                    
                    read_text = f"手順{i}。{step['title']}。{step['text']}"
                    audio_bytes = generate_audio_bytes(read_text)
                    if audio_bytes:
                        st.audio(audio_bytes, format='audio/mp3')
                st.divider()

        excel_data = create_excel_file(steps, manual_number, author_name, create_date, temp_filename)
        st.download_button(
            label="📥 最終版Excelをダウンロード",
            data=excel_data,
            file_name=f"{manual_number}_manual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
