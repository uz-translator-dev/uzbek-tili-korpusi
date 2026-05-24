import os
import re
import pandas as pd
import streamlit as st
from docx import Document

# Sahifa sozlamalari
st.set_page_config(page_title="O'ZBEK TILI KORPUSI", layout="wide")

# Brauzerlarda oq oyna muammosini tuzatuvchi xavfsiz CSS
st.markdown("""
<style>
    /* Umumiy fon */
    .stApp { background-color: #F0F7FF !important; }
    
    /* Sarlavhalar */
    .main-header {
        font-family: 'Segoe UI', sans-serif;
        font-size: 46px;
        color: #1E3A8A !important;
        font-weight: 800;
        text-align: center;
        margin-top: 30px;
    }
    .search-title {
        font-family: 'Georgia', serif;
        font-size: 24px;
        color: #334155 !important;
        font-style: italic;
        text-align: center;
        margin-bottom: 40px;
    }
    
    /* Bosh sahifadagi och yashil bloklar */
    .corpus-box {
        background-color: #D1FAE5 !important;
        border: 2px solid #10B981 !important;
        border-radius: 12px;
        padding: 30px 20px;
        text-align: center;
        margin-bottom: 15px;
    }
    .card-title { color: #065F46 !important; font-size: 22px; font-weight: bold; }
    .card-stat { color: #4B5563 !important; font-size: 14px; }
    .card-desc { color: #047857 !important; font-size: 15px; font-style: italic; }
    
    /* Qidiruv natijalari dizayni (Skrinshotga 100% mos) */
    .custom-sentence-box {
        background-color: #FFFFFF !important;
        padding: 20px 25px;
        border-radius: 6px 6px 0px 0px;
        border-left: 5px solid #64748B !important;
        border-top: 1px solid #E2E8F0 !important;
        border-right: 1px solid #E2E8F0 !important;
        margin-top: 15px;
        font-size: 17px;
        color: #1E293B !important;
        line-height: 1.6;
    }
    
    /* Sariq highlight */
    .highlight {
        background-color: #FDE047 !important;
        color: #000000 !important;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
    }
    
    /* Expander blokining ko'rinishini majburlash (Oq oyna xatosini tuzatish) */
    .stExpander {
        background-color: #F8FAFC !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 5px solid #64748B !important;
        border-radius: 0px 0px 6px 6px !important;
        color: #1E293B !important;
    }
</style>
""", unsafe_allow_html=True)

def extract_tags_from_docx(docx_path):
    tags = {}
    if os.path.exists(docx_path):
        try:
            doc = Document(docx_path)
            for row in doc.tables[0].rows:
                if len(row.cells) >= 2:
                    key = row.cells[0].text.strip()
                    value = row.cells[1].text.strip()
                    if key: tags[key] = value
        except: pass
    return tags

@st.cache_data
def load_umumiy_korpus(folder):
    data = []
    txt_d = os.path.join(folder, "txt_files")
    docx_d = os.path.join(folder, "docx_files")
    if os.path.exists(txt_d):
        for i in range(1, 25):
            t_p = os.path.join(txt_d, f"text_{i}.txt")
            d_p = os.path.join(docx_d, f"tag_{i}.docx")
            if os.path.exists(t_p):
                tags = extract_tags_from_docx(d_p)
                with open(t_p, 'r', encoding='utf-8') as f:
                    sentences = re.split(r'(?<=[.!?])\s+', f.read())
                    for s in sentences:
                        if s.strip():
                            row_dict = {"Fayl": f"text_{i}.txt", "Gap": s.strip()}
                            for k, v in tags.items(): row_dict[k] = v
                            data.append(row_dict)
    return pd.DataFrame(data)

if 'page' not in st.session_state: 
    st.session_state.page = 'home'

UMUMIY_FOLDER = "data/umumiy"
DISKURS_FOLDER = "data/diskurs"

# Orqaga qaytish tugmasi
if st.session_state.page != 'home':
    if st.button("⬅️ KORPUSLAR BO'YICHA QIDIRUV (Bosh sahifa)"):
        st.session_state.page = 'home'
        st.rerun()

# --- BOSH SAHIFA ---
if st.session_state.page == 'home':
    st.markdown('<div class="main-header">O\'ZBEK TILI KORPUSI</div>', unsafe_allow_html=True)
    st.markdown('<div class="search-title">KORPUSLAR BO\'YICHA QIDIRUV</div>', unsafe_allow_html=True)
    
    df_temp = load_umumiy_korpus(UMUMIY_FOLDER)
    total_gap = len(df_temp) if not df_temp.empty else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="corpus-box"><div class="card-title">Umumiy korpus</div><div class="card-stat">24 ta matn | {total_gap:,} ta gap</div><div class="card-desc">(24 ta matn, dinamik hajmli)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="go_um", use_container_width=True): st.session_state.page = 'um'; st.rerun()
    with c2:
        st.markdown('<div class="corpus-box"><div class="card-title">Parallel korpus</div><div class="card-stat">O\'zbek-Turk tili</div><div class="card-desc">(Tillararo ulangan)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="go_par", use_container_width=True): st.session_state.page = 'par'; st.rerun()
    with c3:
        st.markdown('<div class="corpus-box"><div class="card-title">Diskurs korpus</div><div class="card-stat">N-gram tahlili</div><div class="card-desc">(Moslik jadvallari)</div></div>', unsafe_allow_html=True)
        if st.button("Kirish", key="go_dis", use_container_width=True): st.session_state.page = 'dis'; st.rerun()

# --- ICHKI BO'LIMLAR ---
elif st.session_state.page == 'um':
    df = load_umumiy_korpus(UMUMIY_FOLDER)
    st.title("📂 Umumiy korpus")
    
    tab1, tab2 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 Chastotali lug'at"])
    with tab1:
        col_inp, col_btn = st.columns([5, 1])
        with col_inp:
            q = st.text_input("So'z kiriting:", placeholder="Masalan: maktab, til...", label_visibility="collapsed")
        with col_btn:
            lupa_tugmasi = st.button("🔍 Qidirish", use_container_width=True)
            
        if q.strip() or lupa_tugmasi:
            word = q.strip()
            if word:
                pat = re.compile(rf"\b(\w*){re.escape(word)}(\w*)\b", re.IGNORECASE)
                res = df[df['Gap'].apply(lambda x: bool(pat.search(x)))]
                
                if not res.empty:
                    st.success(f"🔍 Jami {len(res)} ta mos gap topildi.")
                    
                    # Statistika expander
                    with st.expander("📈 Qidirilgan so'z bo'yicha statistika"):
                        file_counts = res["Fayl"].value_counts().reset_index()
                        file_counts.columns = ["Fayl nomi", "Qo'llanish soni"]
                        st.dataframe(file_counts, use_container_width=True)
                    
                    # Natijalarni skrinshot formatida chiqarish
                    for _, r in res.iterrows():
                        h_sentence = pat.sub(r'<span class="highlight">\g<0></span>', r['Gap'])
                        st.markdown(f'<div class="custom-sentence-box">{h_sentence}</div>', unsafe_allow_html=True)
                        
                        with st.expander("Metama'lumotlar", expanded=False):
                            meta_data = {k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]}
                            if meta_data:
                                mc1, mc2 = st.columns(2)
                                items = list(meta_data.items())
                                mid = len(items) // 2 + len(items) % 2
                                with mc1:
                                    for k, v in items[:mid]: st.write(f"🔹 **{k}:** {v}")
                                with mc2:
                                    for k, v in items[mid:]: st.write(f"🔹 **{k}:** {v}")
                else: st.warning("Topilmadi.")
    with tab2:
        xlsx_path = os.path.join(UMUMIY_FOLDER, "chastota.xlsx")
        if os.path.exists(xlsx_path): st.dataframe(pd.read_excel(xlsx_path), use_container_width=True)

elif st.session_state.page == 'par':
    st.title("🌐 Parallel korpus (O'zbek-Turk)")
    st.components.v1.iframe("https://uzbek-turk-parallel-korpusi-cnzm5cmc3tkccaysyxai5s.streamlit.app/?embed=true", height=800, scrolling=True)

elif st.session_state.page == 'dis':
    st.title("📊 Diskurs korpus")
    moslik_csv = os.path.join(DISKURS_FOLDER, "moslik.csv")
    if os.path.exists(moslik_csv):
        df_mos = pd.read_csv(moslik_csv)
        d_query = st.text_input("O'zbekcha diskurs n-gram kiriting:")
        if d_query:
            result_mos = df_mos[df_mos.astype(str).apply(lambda x: x.str.contains(d_query, case=False)).any(axis=1)]
            if not result_mos.empty: st.dataframe(result_mos, use_container_width=True)
    else: st.warning("Diskurs CSV fayllari topilmadi.")