import os
import re
from collections import Counter
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="O'ZBEK TILI KORPUSI", layout="wide")

# Rasmdagi chiroyli yashil qutilar va interfeys uslublari uchun CSS
st.markdown("""
<style>
    .stApp { background-color: #F0F7FF !important; }
    .main-header { font-family: 'Segoe UI', sans-serif; font-size: 46px; color: #1E3A8A !important; font-weight: 800; text-align: center; margin-top: 20px; }
    .search-title { font-family: 'Georgia', serif; font-size: 24px; color: #334155 !important; font-style: italic; text-align: center; margin-bottom: 40px; }
    .corpus-box { background-color: #D1FAE5 !important; border: 2px solid #10B981 !important; border-radius: 12px; padding: 30px 20px; text-align: center; min-height: 180px; }
    .card-title { color: #065F46 !important; font-size: 24px; font-weight: bold; margin-bottom: 10px; }
    .card-stat { color: #4B5563 !important; font-size: 15px; margin-bottom: 8px; }
    .card-desc { color: #047857 !important; font-size: 16px; font-style: italic; }
    .highlight { background-color: #FDE047 !important; color: #000000 !important; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
    .sentence-container { background-color: #FFFFFF; padding: 15px 20px; border-radius: 6px; border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 10px; font-size: 16px; color: #1E293B; }
    .stat-box { background-color: #EFF6FF; padding: 15px; border-radius: 8px; border: 1px solid #BFDBFE; margin-bottom: 20px; }
</style>
""", unsafe_allow_html=True)

def extract_tags_from_docx(docx_path):
    tags = {}
    if os.path.exists(docx_path):
        try:
            doc = Document(docx_path)
            if doc.tables and len(doc.tables) > 0:
                for row in doc.tables[0].rows:
                    if len(row.cells) >= 2:
                        key = row.cells[0].text.strip()
                        value = row.cells[1].text.strip()
                        if key: tags[key] = value
        except Exception: pass
    return tags

@st.cache_data
def load_korpus_baza(folder_path, file_count, prefix_txt, prefix_docx, ext_txt, ext_docx):
    data = []
    actual_folder = folder_path
    if not os.path.exists(actual_folder):
        base_name = os.path.basename(folder_path)
        if os.path.exists(base_name): actual_folder = base_name
        elif os.path.exists("data") and base_name in os.listdir("data"): actual_folder = os.path.join("data", base_name)
            
    txt_d = os.path.join(actual_folder, "txt_files")
    docx_d = os.path.join(actual_folder, "docx_files")
    
    if os.path.exists(actual_folder):
        for sub in os.listdir(actual_folder):
            if sub.lower() == "txt_files": txt_d = os.path.join(actual_folder, sub)
            if sub.lower() == "docx_files": docx_d = os.path.join(actual_folder, sub)

    if os.path.exists(txt_d):
        for i in range(1, file_count + 1):
            t_p = os.path.join(txt_d, f"{prefix_txt}{i}.{ext_txt}")
            d_p = os.path.join(docx_d, f"{prefix_docx}{i}.{ext_docx}")
            for f in os.listdir(txt_d):
                if f.lower() == f"{prefix_txt}{i}.{ext_txt}".lower(): t_p = os.path.join(txt_d, f)
            if os.path.exists(docx_d):
                for f in os.listdir(docx_d):
                    if f.lower() == f"{prefix_docx}{i}.{ext_docx}".lower(): d_p = os.path.join(docx_d, f)
            
            if os.path.exists(t_p):
                tags = extract_tags_from_docx(d_p)
                try:
                    with open(t_p, 'r', encoding='utf-8') as f:
                        sentences = re.split(r'(?<=[.!?])\s+', f.read())
                        for s in sentences:
                            if s.strip():
                                row_dict = {"Fayl": os.path.basename(t_p), "Gap": s.strip()}
                                for k, v in tags.items(): row_dict[k] = v
                                data.append(row_dict)
                except Exception: pass
    return pd.DataFrame(data)

def analyze_text_statistics(df_corpus):
    if df_corpus.empty: return [], []
    all_text = " ".join(df_corpus['Gap'].astype(str)).lower()
    words = re.findall(r"\b[b-df-hj-np-rt-vxz'aouei‘g‘shch]+", all_text)
    stop_words = {"va", "bilan", "uchun", "ham", "bu", "ki", "shu", "esa", "bor", "u", "men", "sen", "emas", "kabi", "deb", "tushgan", "bo'g'liq", "viloyat", "respublika"}
    filtered_words = [w for w in words if w not in stop_words and len(w) > 2]
    return Counter(filtered_words).most_common(20), Counter([f"{filtered_words[i]} {filtered_words[i+1]}" for i in range(len(filtered_words) - 1)]).most_common(10)

def show_kwic_search(df_src, key_prefix):
    col_inp, col_btn = st.columns([5, 1])
    with col_inp:
        q = st.text_input("So'z yoki o'zakni kiriting:", placeholder="Masalan: bola, matbuot...", key=f"inp_{key_prefix}")
    with col_btn:
        st.write("<br>", unsafe_allow_html=True)
        lupa = st.button("🔍 Qidirish", key=f"btn_{key_prefix}")
        
    if q.strip() or lupa:
        word = q.strip()
        if word:
            pat = re.compile(rf"({re.escape(word)}[b-df-hj-np-rt-vxz'aouei‘g‘shch]*)", re.IGNORECASE)
            res = df_src[df_src['Gap'].apply(lambda x: bool(pat.search(x)))] if not df_src.empty else pd.DataFrame()
            
            if not res.empty:
                jami_marta = 0
                fayllar = set()
                for _, r in res.iterrows():
                    matches = pat.findall(r['Gap'])
                    jami_marta += len(matches)
                    fayllar.add(r['Fayl'])
                
                st.markdown(f"""
                <div class="stat-box">
                    <h4>📊 Qidiruv statistikasi ("{word}"):</h4>
                    <ul>
                        <li><b>Uchrashish soni (Chastota):</b> {jami_marta} marta kelgan</li>
                        <li><b>Matnlar qamrovi:</b> {len(fayllar)} ta alohida faylda ishlatilgan</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
                
                st.write("🔍 **Kontekstlar ro'yxati (KWIC):**")
                for _, r in res.iterrows():
                    highlighted_sentence = pat.sub(r'<span class="highlight">\g<1></span>', r['Gap'])
                    st.markdown(f'<div class="sentence-container">{highlighted_sentence}</div>', unsafe_allow_html=True)
                    with st.expander("Metama'lumotlar"): 
                        st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
            else: st.warning("Korpus bazasidan ushbu so'z topilmadi.")

UMUMIY_FOLDER = "data/umumiy" if os.path.exists("data/umumiy") else "umumiy"
PUBLISTISTIKA_FOLDER = "data/publististika" if os.path.exists("data/publististika") else "publististika"

# Session state orqali navigatsiyani bog'lash
if 'current_page' not in st.session_state:
    st.session_state.current_page = "🏠 Bosh sahifa"

# --- Yon panel navigatsiyasi ---
st.sidebar.markdown("### 🧭 KORPUS NAVIGATSIYASI")
sidebar_page = st.sidebar.radio(
    "Bo'limni tanlang:", 
    ["🏠 Bosh sahifa", "📂 Umumiy korpus", "🌐 Parallel korpus", "✍️ Publististik korpus"],
    index=["🏠 Bosh sahifa", "📂 Umumiy korpus", "🌐 Parallel korpus", "✍️ Publististik korpus"].index(st.session_state.current_page)
)

# Agar sidebardan tanlansa session stateni yangilash
if sidebar_page != st.session_state.current_page:
    st.session_state.current_page = sidebar_page
    st.rerun()

# =========================================================
# 🏠 1. BOSH SAHIFA (Rasmdagi interfeys to'liq tiklandi)
# =========================================================
if st.session_state.current_page == "🏠 Bosh sahifa":
    st.markdown('<div class="main-header">O\'ZBEK TILI KORPUSI</div>', unsafe_allow_html=True)
    st.markdown('<div class="search-title">KORPUSLAR BO\'YICHA QIDIRUV</div>', unsafe_allow_html=True)
    
    df_u_temp = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    total_gap_um = len(df_u_temp) if not df_u_temp.empty else 0
    
    df_p_temp = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    total_gap_pub = len(df_p_temp) if not df_p_temp.empty else 0

    # Yonma-yon chiroyli 3 ta yashil kartochka
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="corpus-box"><div class="card-title">Umumiy korpus</div><div class="card-stat">24 ta matn | {total_gap_um:,} ta gap</div><div class="card-desc">(24 ta matn, dinamik hajmli)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="btn_enter_um", use_container_width=True):
            st.session_state.current_page = "📂 Umumiy korpus"
            st.rerun()
            
    with c2:
        st.markdown('<div class="corpus-box"><div class="card-title">Parallel korpus</div><div class="card-stat">O\'zbek-Turk tili</div><div class="card-desc">(Tillararo ulangan)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="btn_enter_par", use_container_width=True):
            st.session_state.current_page = "🌐 Parallel korpus"
            st.rerun()
            
    with c3:
        st.markdown(f'<div class="corpus-box"><div class="card-title">Diskurs korpus</div><div class="card-stat">N-gram tahlili</div><div class="card-desc">(Moslik jadvallari)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="btn_enter_pub", use_container_width=True):
            st.session_state.current_page = "✍️ Publististik korpus"
            st.rerun()

# =========================================================
# 📂 2. UMUMIY KORPUS
# =========================================================
elif st.session_state.current_page == "📂 Umumiy korpus":
    df = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    st.title("📂 Umumiy korpus")
    t_u1, t_u2 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 Umumiy chastotali lug'at"])
    with t_u1:
        show_kwic_search(df, "umumiy")
    with t_u2:
        xlsx_path = os.path.join(UMUMIY_FOLDER, "chastota.xlsx")
        if not os.path.exists(xlsx_path) and os.path.exists("umumiy/chastota.xlsx"): xlsx_path = "umumiy/chastota.xlsx"
        if os.path.exists(xlsx_path): st.dataframe(pd.read_excel(xlsx_path), use_container_width=True)

# =========================================================
# 🌐 3. PARALLEL KORPUS (Ortga qaytish strilkasi mukammallashtirildi)
# =========================================================
elif st.session_state.current_page == "🌐 Parallel korpus":
    # Bosh sahifaga qaytish uchun strilka tugmasi
    if st.button("⬅️ Bosh sahifaga qaytish", use_container_width=False):
        st.session_state.current_page = "🏠 Bosh sahifa"
        st.rerun()
        
    st.title("🌐 O'zbek-Turk Parallel Korpusi")
    st.components.v1.iframe("https://uzbek-turk-parallel-korpusi-cnzm5cmc3tkccaysyxai5s.streamlit.app/?embed=true", height=800)

# =========================================================
# ✍️ 4. PUBLISTISTIK KORPUS
# =========================================================
elif st.session_state.current_page == "✍️ Publististik korpus":
    df_pub = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    st.title("✍️ Publististik matnlar korpusi")
    t1, t2, t3 = st.tabs(["🔍 KWIC Qidiruv", "📊 Diskurs tahlil", "🏛️ Mafkuraviy tahlil"])
    
    with t1:
        show_kwic_search(df_pub, "pub")
            
    with t2:
        words, bigrams = analyze_text_statistics(df_pub)
        col1, col2 = st.columns(2)
        col1.write("### 📈 Eng ko'p ishlatilgan so'zlar:")
        col1.dataframe(pd.DataFrame(words, columns=["So'z", "Soni"]), use_container_width=True)
        col2.write("### 🔗 Avtomatik kollokatsiyalar:")
        col2.dataframe(pd.DataFrame(bigrams, columns=["Birikma", "Chastota"]), use_container_width=True)
        
    with t3:
        st.info("### 🏛️ 4-bosqich. Ideologik va ijtimoiy ma'nolarni aniqlash\n\n1. **Muallifning pozitsiyasi:** Matndagi sub'ektiv munosabat.\n2. **Ijtimoiy yoki siyosiy qarashlar:** Davr va mafkura tahlili.\n3. **Auditoriyaga ta'sir qilish usullari:** Ritorik vositalar.\n4. **Diskursdagi asosiy g'oya:** Konseptual yadro.")
