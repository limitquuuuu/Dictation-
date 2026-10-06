import streamlit as st
import docx
from gtts import gTTS
import tempfile
import os
import base64
import re
import jieba
import time
import streamlit.components.v1 as components

# 設定頁面標題與佈局
st.set_page_config(page_title="中英文默書小幫手", layout="centered")

st.title("📚 中英文默書小幫手")

# 初始化 Session State 變數
if "words" not in st.session_state:
    st.session_state.words = []
if "current_idx" not in st.session_state:
    st.session_state.current_idx = 0
if "show_answer" not in st.session_state:
    st.session_state.show_answer = False

# --- 側邊欄：設定與檔案上傳 ---
with st.sidebar:
    st.header("⚙️ 設定與檔案上傳")
    
    uploaded_file = st.file_uploader("上傳默書文件 (.docx)", type=["docx"])
    
    mode = st.radio("選擇練習模式", ["家長/聽寫模式 (顯示答案)", "學生自習模式 (隱藏答案)"])
    
    lang_choice = st.selectbox(
        "選擇語言模式",
        ["英文 (English)", "中文 (廣東話/普通話)"]
    )
    
    speed_fast = st.checkbox("正常語速 (預設為慢速)", value=False)
    
    tld_code = 'com'
    if "英文" in lang_choice:
        accent = st.selectbox("選擇英文發音腔調", ["美音 (US)", "英音 (UK)"])
        tld_code = 'co.uk' if "英音" in accent else 'com'

# --- 核心邏輯：讀取 docx 檔案 ---
if uploaded_file is not None:
    doc = docx.Document(uploaded_file)
    extracted_words = []
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        lines = re.split(r'[\n\r]+', text)
        for line in lines:
            cleaned = line.strip()
            if cleaned:
                extracted_words.append(cleaned)

    if extracted_words:
        # 當上傳新檔案時重置詞庫與索引
        if st.session_state.words != extracted_words:
            st.session_state.words = extracted_words
            st.session_state.current_idx = 0
            st.session_state.show_answer = False
            st.rerun()

# --- 主要內容顯示區 ---
if st.session_state.words:
    total_words = len(st.session_state.words)
    curr_idx = st.session_state.current_idx
    current_word = st.session_state.words[curr_idx]

    # 進度條與狀態顯示
    st.progress((curr_idx + 1) / total_words)
    st.caption(f"進度：第 {curr_idx + 1} / {total_words} 題")

    # 題目卡片顯示（依據模式顯示或遮蔽答案）
    if "家長" in mode or st.session_state.show_answer:
        st.markdown(f"### 🎯 當前題目：**{current_word}**")
    else:
        st.markdown("### 🎯 當前題目：**🙈 [答案已隱藏]**")
        if st.button("👁️ 顯示/揭曉答案"):
            st.session_state.show_answer = True
            st.rerun()

    # 判斷語言代碼
    lang_code = 'en' if "英文" in lang_choice else 'zh-CN'
    full_speech_text = current_word

    # 🎯 針對 iOS Safari 優化的 HTML5 音訊控制元件
    try:
        temp_dir = tempfile.gettempdir()
        audio_path = os.path.join(temp_dir, "dictation_current.mp3")
        
        if lang_code == 'en':
            tts = gTTS(text=full_speech_text, lang=lang_code, tld=tld_code, slow=not speed_fast)
        else:
            tts = gTTS(text=full_speech_text, lang=lang_code, slow=not speed_fast)
            
        tts.save(audio_path)
        
        with open(audio_path, "rb") as f:
            audio_bytes = f.read()
            b64_audio = base64.b64encode(audio_bytes).decode()

        # 生成獨立唯一 ID 供 JS 識別 DOM
        unique_id = f"{curr_idx}_{int(time.time()*1000)}"

        html_code = f"""
        <div style="width: 100%; text-align: center;">
            <audio id="audio_{unique_id}" controls playsinline webkit-playsinline style="width: 100%;">
                <source src="data:audio/mp3;base64,{b64_audio}" type="audio/mp3">
            </audio>
            
            <button onclick="playLocalAudio()" style="margin-top: 10px; width: 100%; padding: 12px; background-color: #4CAF50; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer;">
                🔊 重複發音 (Replay)
            </button>
        </div>

        <script>
            function playLocalAudio() {{
                var audio = document.getElementById('audio_{unique_id}');
                if (audio) {{
                    audio.currentTime = 0;
                    audio.play().catch(function(err) {{
                        console.log("Play failed:", err);
                    }});
                }}
            }}

            setTimeout(function() {{
                playLocalAudio();
            }}, 300);
        </script>
        """
        # 注意：此處已移除有問題的 key 參數
        components.html(html_code, height=120)

    except Exception as e:
        st.error(f"語音生成失敗：{e}")

    st.write("---")

    # 導航按鈕 (上一個 / 下一個)
    col1, col2 = st.columns(2)

    with col1:
        if st.button("⬅️ 上一個", use_container_width=True, disabled=(curr_idx == 0)):
            st.session_state.current_idx -= 1
            st.session_state.show_answer = False
            st.rerun()

    with col2:
        if st.button("下一個 ➡️", use_container_width=True, disabled=(curr_idx == total_words - 1)):
            st.session_state.current_idx += 1
            st.session_state.show_answer = False
            st.rerun()

else:
    st.info("👈 請先於左側邊欄上傳 `.docx` 默書文件檔以開始使用！")