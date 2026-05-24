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

# 2. Korpus yuklash bazasi (Xavfsiz va moslashuvchan)
@st.cache_data
def load_korpus_baza(folder_path, file_count, prefix_txt, prefix_docx, ext_txt, ext_docx):
    data = []
    
    actual_folder = folder_path
    if not os.path.exists(actual_folder):
        base_name = os.path.basename(folder_path)
        if os.path.exists(base_name):
            actual_folder = base_name
        elif os.path.exists("data") and base_name in os.listdir("data"):
            actual_folder = os.path.join("data", base_name)
            
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

# Papka manzillari tekshiruvi
UMUMIY_FOLDER = "data/umumiy" if os.path.exists("data/umumiy") else ("umumiy" if os.path.exists("umumiy") else "data/umumiy")
PUBLISTISTIKA_FOLDER = "data/publististika" if os.path.exists("data/publististika") else ("publististika" if os.path.exists("publististika") else "data/publististika")

# --- Navigatsiya Paneli ---
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
        st.markdown(f'<div class="corpus-box"><div class="card-title">Publististik matnlar korpusi</div><div class="card-stat">21 ta matn | {total_gap_pub:,} ta gap</div><div class="card-desc">(Diskursiv tahlil moduli)</div></div>', unsafe_allow_html=True)

    # Diagnostika paneli
    st.write("---")
    st.subheader("🛠️ Tizim fayllari diagnostikasi:")
    base_to_check = "data" if os.path.exists("data") else "."
    found_folders = os.listdir(base_to_check)
    st.write(f"📂 Mavjud papkalar/fayllar: `{found_folders}`")
    
    p_check = None
    for item in found_folders:
        if item.lower() == "publististika": p_check = os.path.join(base_to_check, item)
    if base_to_check == "." and "publististika" in found_folders: p_check = "publististika"
        
    if p_check and os.path.exists(p_check):
        st.success(f"✅ Publististika papkasi topildi: `{p_check}`")
        subs = os.listdir(p_check)
        t_sub = None
        for s in subs:
            if s.lower() == "txt_files": t_sub = os.path.join(p_check, s)
        if t_sub and os.path.exists(t_sub):
            f_list = os.listdir(t_sub)
            st.success(f"✅ `txt_files` ichida {len(f_list)} ta fayl mavjud.")
        else:
            st.error("❌ Xato: Publististika ichida `txt_files` topilmadi!")
    else:
        st.error("❌ Xato: Publististika papkasi topilmadi. GitHub'ga to'g'ri yuklanganini tekshiring!")

# =========================================================
# 📂 2. UMUMIY KORPUS
# =========================================================
elif page_selection == "📂 Umumiy korpus":
    df = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    st.title("📂 Umumiy korpus")
    
    tab1, tab2 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 Chastotali lug'at"])
    with tab1:
        col_inp, col_btn = st.columns([5, 1])
        with col_inp: q = st.text_input("So'z kiriting:", placeholder="Masalan: maktab, til...", label_visibility="collapsed")
        with col_btn: lupa_tugmasi = st.button("🔍 Qidirish", use_container_width=True)
        
        if q.strip() or lupa_tugmasi:
            word = q.strip()
            if word:
                pat = re.compile(rf"({re.escape(word)})", re.IGNORECASE)
                res = df[df['Gap'].apply(lambda x: bool(pat.search(x)))] if not df.empty else pd.DataFrame()
                if not res.empty:
                    st.success(f"🔍 Jami {len(res)} ta mos gap topildi.")
                    for _, r in res.iterrows():
                        h_sentence = pat.sub(r'<span class="highlight">\g<1></span>', r['Gap'])
                        st.markdown(f'<div class="custom-sentence-box">{h_sentence}</div>', unsafe_allow_html=True)
                        with st.expander("Metama'lumotlar"): st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
                else: st.warning("Topilmadi.")
    with tab2:
        xlsx_path = os.path.join(UMUMIY_FOLDER, "chastota.xlsx")
        if not os.path.exists(xlsx_path) and os.path.exists("umumiy/chastota.xlsx"): xlsx_path = "umumiy/chastota.xlsx"
        if os.path.exists(xlsx_path): st.dataframe(pd.read_excel(xlsx_path), use_container_width=True)

# =========================================================
# 🌐 3. PARALLEL KORPUS
# =========================================================
elif page_selection == "🌐 Parallel korpus":
    st.markdown('<div class="inner-header"><h2>🌐 O‘zbek-Turk Parallel Korpusi</h2><p>Tillararo qidiruv tizimi</p></div>', unsafe_allow_html=True)
    st.info("ℹ️ Parallel korpus tizimi quyidagi oynada ochiladi. Bosh sahifaga qaytish uchun chap paneldagi navigatsiyadan foydalaning.")
    st.components.v1.iframe("https://uzbek-turk-parallel-korpusi-cnzm5cmc3tkccaysyxai5s.streamlit.app/?embed=true", height=800, scrolling=True)

# =========================================================
# ✍️ 4. PUBLISTISTIK MATNLAR KORPUSI
# =========================================================
elif page_selection == "✍️ Publististik matnlar korpusi":
    df_pub = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    st.title("✍️ Publististik matnlar korpusi (50 000 ta so'z)")
    
    tab_p1, tab_p2, tab_p3 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 3-bosqich. Diskurs tahlili", "🏛️ 4-bosqich. Ideologik va ijtimoiy ma'nolarni aniqlash"])
    
    with tab_p1:
        col_inp, col_btn = st.columns([5, 1])
        with col_inp: q_pub = st.text_input("Publististik korpusdan so'z kiriting:", placeholder="Masalan: matbuot, xalq...", label_visibility="collapsed")
        with col_btn: lupa_pub = st.button("🔍 Qidirish", key="l_pub", use_container_width=True)
        
        if q_pub.strip() or lupa_pub:
            word = q_pub.strip()
            if word:
                pat = re.compile(rf"({re.escape(word)})", re.IGNORECASE)
                res = df_pub[df_pub['Gap'].apply(lambda x: bool(pat.search(x)))] if not df_pub.empty else pd.DataFrame()
                if not res.empty:
                    st.success(f"🔍 Publististik korpus bo'yicha jami {len(res)} ta mos gap aniqlandi.")
                    for _, r in res.iterrows():
                        h_sentence = pat.sub(r'<span class="highlight">\g<1></span>', r['Gap'])
                        st.markdown(f'<div class="custom-sentence-box">{h_sentence}</div>', unsafe_allow_html=True)
                        with st.expander("Metama'lumotlar"): st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
                else: st.warning("Publististik korpus matnlari ichida ushbu so'z topilmadi.")

    with tab_p2:
        st.subheader("📋 Talaba korpus asosida quyidagi jihatlarni tahlil qiladi:")
        col_an1, col_an2 = st.columns(2)
        with col_an1:
            st.markdown("""<div class="analysis-box"><h3 style="color:#1E3A8A; margin-top:0;">1. Nutq strategiyalari</h3><p><b>Quyidagilar aniqlanadi:</b></p><ul><li>🎯 Persuaziv strategiyalar</li><li>💎 Baholovchi birliklar</li><li>🔥 Emotsional ifodalar</li><li>🧩 Argumentatsiya usullari</li></ul></div>""", unsafe_allow_html=True)
        with col_an2:
            st.markdown("""<div class="analysis-box"><h3 style="color:#1E3A8A; margin-top:0