import os
import re
import pandas as pd
import streamlit as st
from docx import Document

# Sahifa sozlamalari
st.set_page_config(page_title="O'ZBEK TILI KORPUSI", layout="wide")

# Vizual uslublar (Dizayn talablaringiz asosida)
st.markdown("""
<style>
    .stApp { background-color: #F0F7FF !important; }
    .main-header { font-family: 'Segoe UI', sans-serif; font-size: 46px; color: #1E3A8A !important; font-weight: 800; text-align: center; margin-top: 10px; }
    .search-title { font-family: 'Georgia', serif; font-size: 24px; color: #334155 !important; font-style: italic; text-align: center; margin-bottom: 30px; }
    .corpus-box { background-color: #D1FAE5 !important; border: 2px solid #10B981 !important; border-radius: 12px; padding: 25px 15px; text-align: center; margin-bottom: 15px; }
    .card-title { color: #065F46 !important; font-size: 22px; font-weight: bold; }
    .card-stat { color: #4B5563 !important; font-size: 14px; }
    .card-desc { color: #047857 !important; font-size: 15px; font-style: italic; }
    .inner-header { background-color: #1E293B; color: #FFFFFF; padding: 20px; border-radius: 8px; margin-bottom: 25px; text-align: center; }
    .highlight { background-color: #FDE047 !important; color: #000000 !important; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
    .custom-sentence-box { background-color: #FFFFFF !important; padding: 20px 25px; border-radius: 6px 6px 0px 0px; border-left: 5px solid #64748B !important; border-top: 1px solid #E2E8F0 !important; border-right: 1px solid #E2E8F0 !important; margin-top: 15px; font-size: 17px; color: #1E293B !important; line-height: 1.6; }
    .stExpander { border-radius: 0px 0px 6px 6px !important; border-left: 5px solid #64748B !important; background-color: #F8FAFC !important; margin-top: -1px !important; }
    .analysis-box { background-color: #FFFFFF; padding: 20px; border-radius: 8px; border-top: 4px solid #1E3A8A; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px; }
</style>
""", unsafe_allow_html=True)

# 1. DOCX teglarini xavfsiz o'qish
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
        except Exception:
            pass
    return tags

# 2. Korpus yuklash bazasi (Katta-kichik harflar va yo'llardan himoyalangan)
@st.cache_data
def load_korpus_baza(folder_path, file_count, prefix_txt, prefix_docx, ext_txt, ext_docx):
    data = []
    
    # Papka mavjudligini aniqlash (tashqarida yoki data ichida bo'lishiga qarab)
    actual_folder = folder_path
    if not os.path.exists(actual_folder):
        base_name = os.path.basename(folder_path)
        if os.path.exists(base_name):
            actual_folder = base_name
        elif os.path.exists("data") and base_name in os.listdir("data"):
            actual_folder = os.path.join("data", base_name)
            
    # Ichki papkalarni katta-kichik harfga nisbatan qidirish
    txt_d = os.path.join(actual_folder, "txt_files")
    docx_d = os.path.join(actual_folder, "docx_files")
    
    if os.path.exists(actual_folder):
        for sub in os.listdir(actual_folder):
            if sub.lower() == "txt_files": txt_d = os.path.join(actual_folder, sub)
            if sub.lower() == "docx_files": docx_d = os.path.join(actual_folder, sub)

    if os.path.exists(txt_d):
        for i in range(1, file_count + 1):
            t_file_name = f"{prefix_txt}{i}.{ext_txt}"
            d_file_name = f"{prefix_docx}{i}.{ext_docx}"
            
            t_p = os.path.join(txt_d, t_file_name)
            d_p = os.path.join(docx_d, d_file_name)
            
            # Linux serverda harflar o'zgarib qolgan bo'lsa xavfsiz qidirish
            for f in os.listdir(txt_d):
                if f.lower() == t_file_name.lower(): t_p = os.path.join(txt_d, f)
            if os.path.exists(docx_d):
                for f in os.listdir(docx_d):
                    if f.lower() == d_file_name.lower(): d_p = os.path.join(docx_d, f)
            
            if os.path.exists(t_p):
                tags = extract_tags_from_docx(d_p)
                try:
                    with open(t_p, 'r', encoding='utf-8') as f:
                        text_content = f.read()
                        sentences = re.split(r'(?<=[.!?])\s+', text_content)
                        for s in sentences:
                            if s.strip():
                                row_dict = {"Fayl": os.path.basename(t_p), "Gap": s.strip()}
                                for k, v in tags.items(): row_dict[k] = v
                                data.append(row_dict)
                except Exception:
                    pass
    return pd.DataFrame(data)

# --- Aqlli Yo'l Tanlagich (GitHub'ga qanday usulda yuklanishidan qat'i nazar ishlaydi) ---
UMUMIY_FOLDER = "data/umumiy" if os.path.exists("data/umumiy") else ("umumiy" if os.path.exists("umumiy") else "data/umumiy")
PUBLISTISTIKA_FOLDER = "data/publististika" if os.path.exists("data/publististika") else ("publististika" if os.path.exists("publististika") else "data/publististika")

# --- 🧭 NAVIGATSIYA PANEL (SIDEBAR) ---
st.sidebar.markdown("### 🧭 KORPUS NAVIGATSIYASI")
page_selection = st.sidebar.radio(
    "Bo'limni tanlang:",
    ["🏠 Bosh sahifa", "📂 Umumiy korpus", "🌐 Parallel korpus", "✍️ Publististik matnlar korpusi"]
)

# =========================================================
# 🏠 1. BOSH SAHIFA
# =========================================================
if page_selection == "🏠 Bosh sahifa":
    st.markdown('<div class="main-header">O\'ZBEK TILI KORPUSI</div>', unsafe_allow_html=True)
    st.markdown('<div class="search-title">KORPUSLAR BO\'YICHA QIDIRUV</div>', unsafe_allow_html=True)
    
    df_um_temp = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    total_gap_um = len(df_um_temp) if not df_um_temp.empty else 0
    
    df_pub_temp = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    total_gap_pub = len(df_pub_temp) if not df_pub_temp.empty else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="corpus-box"><div class="card-title">Umumiy korpus</div><div class="card-stat">24 ta matn | {total_gap_um:,} ta gap</div><div class="card-desc">(24 ta matn, dinamik hajmli)</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="corpus-box"><div class="card-title">Parallel korpus</div><div class="card-stat">O\'zbek-Turk tili</div><div class="card-desc">(Tillararo ulangan tizim)</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class