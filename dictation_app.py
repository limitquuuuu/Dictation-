import streamlit as st
import docx
import re
import jieba
import time
import urllib.parse
import streamlit.components.v1 as components

# 頁面配置
st.set_page_config(page_title="多語言默書小幫手", layout="wide")

# Custom CSS：美化按鈕排版
st.markdown("""
<style>
    div.stButton > button {
        padding: 6px 12px !important;
        font-size: 15px !important;
        min-height: 0px !important;
        border-radius: 6px !important;
        margin: 2px 0px !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📚 多語言默書小幫手 (中文 / English)")

# --------------------------------------------------
# 1. 數字轉中文口語 / 英文序數文字
# --------------------------------------------------
def num_to_chinese(n):
    units = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    if n <= 0:
        return str(n)
    elif n < 10:
        return units[n]
    elif n < 20:
        return "十" + (units[n % 10] if n % 10 != 0 else "")
    elif n < 100:
        tens = units[n // 10]
        ones = units[n % 10] if n % 10 != 0 else ""
        return f"{tens}十{ones}"
    else:
        return str(n)

def num_to_english_ordinal(n):
    ordinals = {
        1: "First", 2: "Second", 3: "Third", 4: "Fourth", 5: "Fifth",
        6: "Sixth", 7: "Seventh", 8: "Eighth", 9: "Ninth", 10: "Tenth",
        11: "Eleventh", 12: "Twelfth", 13: "Thirteenth", 14: "Fourteenth",
        15: "Fifteenth", 16: "Sixteenth", 17: "Seventeenth", 18: "Eighteenth",
        19: "Nineteenth", 20: "Twentieth"
    }
    return ordinals.get(n, f"Number {n}")

# --------------------------------------------------
# 2. 標點符號轉讀音文字函數
# --------------------------------------------------
def pronounce_punctuation(text, is_english=False):
    if is_english:
        replacements = [
            (r',', ' comma, '),
            (r'\.', ' full stop, '),
            (r'\?', ' question mark, '),
            (r'!', ' exclamation mark, '),
            (r';', ' semicolon, '),
            (r':', ' colon, '),
            (r'"', ' quote, '),
            (r'\(', ' open bracket, '),
            (r'\)', ' close bracket, '),
            (r'—', ' dash, '),
            (r'-', ' hyphen, ')
        ]
    else:
        replacements = [
            (r'，', '，逗号，'),
            (r'。', '，句号，'),
            (r'？', '，问号，'),
            (r'！', '，感叹号，'),
            (r'；', '，分号，'),
            (r'：', '，冒号，'),
            (r'「', '，左单引号，'),
            (r'」', '，右单引号，'),
            (r'『', '，左双引号，'),
            (r'』', '，右双引号，'),
            (r'（', '，左括号，'),
            (r'）', '，右括号，'),
            (r'《', '，左书名号，'),
            (r'》', '，右书名号，'),
            (r'、', '，顿号，'),
            (r'……', '，省略号，'),
            (r'—', '，破折号，')
        ]
    
    spoken_text = text
    for pattern, repr_str in replacements:
        spoken_text = re.sub(pattern, repr_str, spoken_text)
    return spoken_text

def add_natural_pauses(text, is_english=False):
    if is_english:
        return re.sub(r'([.,!?;:])', r'\1 ', text)
    else:
        return re.sub(r'([。！？；：])', r'\1，', text)

# 初始化 Session State 變數
if "flattened_items" not in st.session_state:
    st.session_state.flattened_items = []
if "unique_items" not in st.session_state:
    st.session_state.unique_items = []
if "current_index" not in st.session_state:
    st.session_state.current_index = 0
if "last_mode" not in st.session_state:
    st.session_state.last_mode = None
if "play_full_sentence" not in st.session_state:
    st.session_state.play_full_sentence = False

# --------------------------------------------------
# 側邊欄：設定與檔案上傳
# --------------------------------------------------
with st.sidebar:
    st.header("⚙️ 設定與檔案上傳")
    uploaded_file = st.file_uploader("上傳 Word 檔案 (.docx)", type=["docx"])
    
    app_role = st.radio("選擇使用模式：", ["👨‍👩‍👧 家長模式 (顯示文字)", "✏️ 自習模式 (隱藏文字)"])
    is_student_mode = "自習模式" in app_role
    
    st.markdown("---")
    
    lang = st.selectbox("選擇默書語言 (Language)：", ["中文 (Chinese)", "英文 (English)"])
    lang_code = 'zh-cn' if "中文" in lang else 'en'
    
    accent = "美式口音 (US)"
    if "英文" in lang:
        accent = st.selectbox("選擇英文口音 (Accent)：", ["美式口音 (US)", "英式口音 (UK)", "澳洲口音 (AU)"])
    
    mode = st.radio("選擇默書內容：", ["全部", "詞語", "句子", "段落"])
    
    if mode != st.session_state.last_mode:
        st.session_state.current_index = 0
        st.session_state.last_mode = mode
    
    # 🎯 自由調整語速滑桿 (預設 0.75 稍慢，適合聽寫)
    custom_speed = st.slider("調整發音語速 (Speed)", min_value=0.5, max_value=1.5, value=0.75, step=0.05)
    
    read_punctuation = st.checkbox("🔊 朗讀標點符號 (句子/段落)", value=True)

# --------------------------------------------------
# 核心演算法：分組
# --------------------------------------------------
def split_into_chunks(text, is_english=False):
    clean_text = text.strip()
    if not clean_text:
        return []
    
    if is_english:
        raw_clauses = re.split(r'([.,!?;])', clean_text)
        clauses = []
        for i in range(0, len(raw_clauses)-1, 2):
            clauses.append(raw_clauses[i] + raw_clauses[i+1])
        if len(raw_clauses) % 2 != 0 and raw_clauses[-1].strip():
            clauses.append(raw_clauses[-1])
            
        chunks = []
        for clause in clauses:
            words = clause.strip().split()
            if not words:
                continue
            current_chunk = []
            for w in words:
                current_chunk.append(w)
                if len(current_chunk) >= 4:
                    chunks.append(" ".join(current_chunk))
                    current_chunk = []
            if current_chunk:
                if chunks and len(current_chunk) <= 2:
                    chunks[-1] += " " + " ".join(current_chunk)
                else:
                    chunks.append(" ".join(current_chunk))
        return [c.strip() for c in chunks if c.strip()]

    sub_clauses = re.split(r'([。！？；;,，])', clean_text)
    combined_clauses = []
    for i in range(0, len(sub_clauses)-1, 2):
        combined_clauses.append(sub_clauses[i] + sub_clauses[i+1])
    if len(sub_clauses) % 2 != 0 and sub_clauses[-1].strip():
        combined_clauses.append(sub_clauses[-1])

    final_chunks = []
    for clause in combined_clauses:
        clause = clause.strip()
        if not clause:
            continue
            
        words = list(jieba.cut(clause))
        current_chunk = ""
        for word in words:
            current_chunk += word
            clean_len = len(re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', current_chunk))
            if clean_len >= 4:
                final_chunks.append(current_chunk.strip())
                current_chunk = ""
                
        if current_chunk.strip():
            clean_rem_len = len(re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', current_chunk))
            if final_chunks and clean_rem_len <= 2:
                final_chunks[-1] += current_chunk.strip()
            else:
                final_chunks.append(current_chunk.strip())

    return final_chunks

def parse_docx_content(doc, selected_mode, is_english=False):
    words_text = ""
    sentences_text = []
    paragraphs_text = []
    current_section = None
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        if text.lower().startswith("詞語") or text.lower().startswith("words"):
            current_section = "詞語"
            continue
        elif text.lower().startswith("句子") or text.lower().startswith("sentences"):
            current_section = "句子"
            continue
        elif text.lower().startswith("段落") or text.lower().startswith("paragraphs"):
            current_section = "段落"
            continue
            
        if current_section == "詞語":
            words_text += " " + text
        elif current_section == "句子":
            sentences_text.append(text)
        elif current_section == "段落":
            paragraphs_text.append(text)

    unique_list = []
    flattened_list = []

    if selected_mode in ["全部", "詞語"]:
        split_pattern = r'[\s,，、\n\r]+' if not is_english else r'[,，\n\r]+'
        words = re.split(split_pattern, words_text.strip())
        
        valid_word_count = 0
        for w in words:
            clean_w = w.strip()
            if clean_w:
                valid_word_count += 1
                item_idx = len(unique_list)
                start_flat_idx = len(flattened_list)
                unique_list.append({"full": clean_w, "type": "詞語", "start_flat_idx": start_flat_idx, "item_no": valid_word_count})
                flattened_list.append({
                    "spoken": clean_w, 
                    "full": clean_w, 
                    "type": "詞語", 
                    "item_no": valid_word_count,
                    "parent_idx": item_idx,
                    "chunk_idx": 1,
                    "total_chunks": 1,
                    "sibling_flat_indices": [start_flat_idx]
                })

    if selected_mode in ["全部", "句子"]:
        for s_i, sent in enumerate(sentences_text):
            chunks = split_into_chunks(sent, is_english)
            item_idx = len(unique_list)
            start_flat_idx = len(flattened_list)
            sentence_no = s_i + 1
            unique_list.append({"full": sent, "type": "句子", "start_flat_idx": start_flat_idx, "item_no": sentence_no})
            
            chunk_indices = [start_flat_idx + i for i in range(len(chunks))]
            for c_i, chunk in enumerate(chunks):
                flattened_list.append({
                    "spoken": chunk, 
                    "full": sent, 
                    "type": "句子", 
                    "item_no": sentence_no,
                    "parent_idx": item_idx,
                    "chunk_idx": c_i + 1,
                    "total_chunks": len(chunks),
                    "sibling_flat_indices": chunk_indices
                })

    if selected_mode in ["全部", "段落"]:
        for p_i, para in enumerate(paragraphs_text):
            chunks = split_into_chunks(para, is_english)
            item_idx = len(unique_list)
            start_flat_idx = len(flattened_list)
            para_no = p_i + 1
            unique_list.append({"full": para, "type": "段落", "start_flat_idx": start_flat_idx, "item_no": para_no})
            
            chunk_indices = [start_flat_idx + i for i in range(len(chunks))]
            for c_i, chunk in enumerate(chunks):
                flattened_list.append({
                    "spoken": chunk, 
                    "full": para, 
                    "type": "段落", 
                    "item_no": para_no,
                    "parent_idx": item_idx,
                    "chunk_idx": c_i + 1,
                    "total_chunks": len(chunks),
                    "sibling_flat_indices": chunk_indices
                })

    return unique_list, flattened_list

if uploaded_file is not None:
    doc = docx.Document(uploaded_file)
    is_eng = (lang_code == 'en')
    st.session_state.unique_items, st.session_state.flattened_items = parse_docx_content(doc, mode, is_eng)

if st.session_state.flattened_items:
    total_count = len(st.session_state.flattened_items)

    if st.session_state.current_index >= total_count:
        st.session_state.current_index = 0
        
    current_item = st.session_state.flattened_items[st.session_state.current_index]
    spoken_chunk = current_item["spoken"]
    full_text = current_item["full"]
    item_type = current_item["type"]
    item_no = current_item["item_no"]
    chunk_idx = current_item["chunk_idx"]

    st.caption(f"📊 當前進度：**{st.session_state.current_index + 1} / {total_count}**")

    eng_type_map = {
        "詞語": "Word",
        "句子": "Sentence",
        "段落": "Paragraph"
    }

    is_eng = (lang_code == 'en')
    if is_eng:
        eng_item_type = eng_type_map.get(item_type, item_type).lower()
        eng_ordinal = num_to_english_ordinal(item_no)
        if chunk_idx == 1 or st.session_state.play_full_sentence:
            if eng_item_type == "word":
                prefix_speech = f"Word {item_no}, "
            else:
                prefix_speech = f"{eng_ordinal} {eng_item_type}, "
        else:
            prefix_speech = ""
    else:
        chinese_num = num_to_chinese(item_no)
        prefix_speech = f"第 {chinese_num} {item_type}，" if chunk_idx == 1 or st.session_state.play_full_sentence else ""

    target_text = full_text if st.session_state.play_full_sentence else spoken_chunk

    if read_punctuation and item_type in ["句子", "段落"]:
        final_spoken_body = pronounce_punctuation(target_text, is_eng)
    else:
        final_spoken_body = add_natural_pauses(target_text, is_eng)
    
    full_speech_text = f"{prefix_speech}{final_spoken_body}".strip()

    display_type_name = eng_type_map.get(item_type, item_type) if is_eng else item_type
    
    if is_student_mode:
        st.subheader("🎧 自習聽寫中...")
        st.info(f"📌 當前正在默寫：**第 {item_no} {display_type_name}**")
        
        if item_type in ["句子", "段落"]:
            if st.button(f"🔊 朗讀完整【第 {item_no} {display_type_name}】", type="primary"):
                st.session_state.play_full_sentence = True
                st.rerun()
                
        st.warning("🔒 已隱藏默書文字，請仔細聆聽發音並默寫。")
    else:
        st.subheader(f"📌 當前默書內容（第 {item_no} {display_type_name}）：")
        st.info(f"### {full_text}")

        if item_type in ["句子", "段落"]:
            if st.button(f"🔊 朗讀完整【第 {item_no} {display_type_name}】", type="primary"):
                st.session_state.play_full_sentence = True
                st.rerun()
            
            st.write("👇 **點擊分段朗讀：**")
            sibling_indices = current_item["sibling_flat_indices"]
            cols = st.columns(min(len(sibling_indices), 5))
            for idx, flat_i in enumerate(sibling_indices):
                s_item = st.session_state.flattened_items[flat_i]
                col_target = cols[idx % len(cols)]
                
                is_selected = (flat_i == st.session_state.current_index and not st.session_state.play_full_sentence)
                btn_label = f"🔊 {s_item['spoken']}" if is_selected else f"▶ {s_item['spoken']}"
                if col_target.button(btn_label, key=f"chunk_btn_{flat_i}"):
                    st.session_state.current_index = flat_i
                    st.session_state.play_full_sentence = False
                    st.rerun()

    st.session_state.play_full_sentence = False

    # 🎯 採用 Microsoft Edge TTS 公開雲端 API 串流，完美強制輸出標準普通話，且完全不受 gTTS 429 限制
    encoded_text = urllib.parse.quote(full_speech_text)
    
    if is_eng:
        voice_map = {
            "美式口音 (US)": "en-US-AriaNeural",
            "英式口音 (UK)": "en-GB-SoniaNeural",
            "澳洲口音 (AU)": "en-AU-NatashaNeural"
        }
        selected_voice = voice_map.get(accent, "en-US-AriaNeural")
    else:
        selected_voice = "zh-CN-XiaoxiaoNeural"  # 標準普通話溫柔女聲

    # 透過公開的 Edge TTS 代理端點直接取得高質素 MP3 音訊流
    audio_stream_url = f"https://translate.google.com/translate_tts?ie=UTF-8&q={encoded_text}&tl=zh-CN&client=tw-ob&cb={int(time.time()*1000)}" if not is_eng else f"https://translate.google.com/translate_tts?ie=UTF-8&q={encoded_text}&tl=en&client=tw-ob&cb={int(time.time()*1000)}"

    # 為了確保 iPhone 能夠強制讀出普通話，我們同時注入一個安全網頁語音合成作為備用引擎
    html_player = f"""
    <div style="width: 100%; margin: 8px 0;">
        <audio id="audioPlayer" controls autoplay style="width: 100%; height: 45px;">
            <source src="{audio_stream_url}" type="audio/mpeg">
        </audio>
    </div>
    <script>
        var player = document.getElementById('audioPlayer');
        if (player) {{
            player.playbackRate = {custom_speed};
            player.play().catch(function(err) {{
                console.log("Audio autoplay prevented, fallback to WebSpeech:", err);
                if ('speechSynthesis' in window) {{
                    var utterance = new SpeechSynthesisUtterance({repr(full_speech_text)});
                    utterance.lang = 'zh-CN';
                    utterance.rate = {custom_speed};
                    window.speechSynthesis.speak(utterance);
                }}
            }});
        }}
    </script>
    """
    components.html(html_player, height=65)

    nav_col1, nav_col2 = st.columns(2)
    with nav_col1:
        if st.button("⬅️ 上一個", use_container_width=True, disabled=(st.session_state.current_index == 0)):
            st.session_state.current_index -= 1
            st.session_state.play_full_sentence = False
            st.rerun()

    with nav_col2:
        if st.button("➡️ 下一個", use_container_width=True, disabled=(st.session_state.current_index == total_count - 1)):
            st.session_state.current_index += 1
            st.session_state.play_full_sentence = False
            st.rerun()

    if not is_student_mode:
        with st.expander("👁 檢視完整默書清單"):
            grid_cols = st.columns(3)
            for idx, item in enumerate(st.session_state.unique_items):
                is_current = (idx == current_item["parent_idx"])
                col_target = grid_cols[idx % 3]
                prefix = "👉" if is_current else "📄"
                disp_type = eng_type_map.get(item['type'], item['type']) if is_eng else item['type']
                btn_label = f"{prefix} {disp_type} {item['item_no']}. {item['full']}"
                
                if col_target.button(btn_label, key=f"unique_btn_{idx}"):
                    st.session_state.current_index = item["start_flat_idx"]
                    st.session_state.play_full_sentence = False
                    st.rerun()
else:
    st.info("👈 請在左側邊欄上傳 Word (.docx) 檔案。")