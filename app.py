import os
import re
from collections import Counter
import pandas as pd
import streamlit as st
from docx import Document

# Sahifa sozlamalari
st.set_page_config(page_title="O'ZBEK TILI KORPUSI", layout="wide")

# Sariq rangni xavfsiz integratsiya qilish uchun CSS
st.markdown("""
<style>
    .highlight {
        background-color: #FDE047 !important;
        color: #000000 !important;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
    }
    .sentence-container {
        background-color: #FFFFFF;
        padding: 15px 20px;
        border-radius: 6px;
        border-left: 5px solid #1E3A8A;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 10px;
        font-size: 16px;
        color: #1E293B;
    }
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

# 2. Korpus yuklash bazasi
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

# 3. Avtomatik statistik tahlil mexanizmi
def analyze_text_statistics(df_corpus):
    if df_corpus.empty:
        return [], []
    
    all_text = " ".join(df_corpus['Gap'].astype(str)).lower()
    words = re.findall(r"\b[b-df-hj-np-rt-vxz'aouei‘g‘shch]+", all_text)
    
    stop_words = {
        "va", "bilan", "uchun", "ham", "bu", "ki", "shu", "esa", "bor", "u", 
        "men", "sen", "emas", "kabi", "deb", "tushgan", "bo'g'liq", "viloyat", "respublika"
    }
    
    filtered_words = [w for w in words if w not in stop_words and len(w) > 2]
    
    word_counts = Counter(filtered_words).most_common(20)
    
    bigrams = []
    for i in range(len(filtered_words) - 1):
        bigrams.append(f"{filtered_words[i]} {filtered_words[i+1]}")
    bigram_counts = Counter(bigrams).most_common(10)
    
    return word_counts, bigram_counts

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
    st.title("O'ZBEK TILI KORPUSI")
    st.subheader("KORPUSLAR BO'YICHA QIDIRUV TIZIMI")
    
    df_um_temp = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    total_gap_um = len(df_um_temp) if not df_um_temp.empty else 0
    
    df_pub_temp = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    total_gap_pub = len(df_pub_temp) if not df_pub_temp.empty else 0

    c1, c2, c3 = st.columns(3)
    with c1: st.success(f"📂 Umumiy korpus\n\n24 ta matn | {total_gap_um:,} ta gap")
    with c2: st.info("🌐 Parallel korpus\n\nO'zbek-Turk tili | Tillararo ulangan")
    with c3: st.warning(f"✍️ Publististik matnlar korpusi\n\n21 ta matn | {total_gap_pub:,} ta gap")

# =========================================================
# 📂 2. UMUMIY KORPUS
# =========================================================
elif page_selection == "📂 Umumiy korpus":
    df = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    st.title("📂 Umumiy korpus")
    
    tab1, tab2 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 Chastotali lug'at"])
    with tab1:
        q = st.text_input("So'z kiriting:", placeholder="Masalan: maktab, til...")
        if q.strip():
            word = q.strip()
            pat = re.compile(rf"({re.escape(word)})", re.IGNORECASE)
            res = df[df['Gap'].apply(lambda x: bool(pat.search(x)))] if not df.empty else pd.DataFrame()
            if not res.empty:
                st.success(f"🔍 Jami {len(res)} ta mos gap topildi.")
                for _, r in res.iterrows():
                    highlighted_sentence = pat.sub(r'<span class="highlight">\g<1></span>', r['Gap'])
                    st.markdown(f'<div class="sentence-container">{highlighted_sentence}</div>', unsafe_allow_html=True)
                    with st.expander("Metama'lumotlar"): 
                        st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
            else: st.warning("Topilmadi.")
    with tab2:
        xlsx_path = os.path.join(UMUMIY_FOLDER, "chastota.xlsx")
        if not os.path.exists(xlsx_path) and os.path.exists("umumiy/chastota.xlsx"): xlsx_path = "umumiy/chastota.xlsx"
        if os.path.exists(xlsx_path): st.dataframe(pd.read_excel(xlsx_path), use_container_width=True)

# =========================================================
# 🌐 3. PARALLEL KORPUS
# =========================================================
elif page_selection == "🌐 Parallel korpus":
    st.title("🌐 O'zbek-Turk Parallel Korpusi")
    st.components.v1.iframe("https://uzbek-turk-parallel-korpusi-cnzm5cmc3tkccaysyxai5s.streamlit.app/?embed=true", height=800, scrolling=True)

# =========================================================
# ✍️ 4. PUBLISTISTIK MATNLAR KORPUSI
# =========================================================
elif page_selection == "✍️ Publististik matnlar korpusi":
    df_pub = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    st.title("✍️ Publististik matnlar korpusi (50 000 ta so'z)")
    
    tab_p1, tab_p2, tab_p3 = st.tabs(["🔍 Kontekstli qidiruv (KWIC)", "📊 3-bosqich. Diskurs tahlili", "🏛️ 4-bosqich. Ideologik va ijtimoiy ma'nolarni aniqlash"])
    
    with tab_p1:
        q_pub = st.text_input("Publististik korpusdan so'z kiriting:", placeholder="Masalan: matbuot, xalq...")
        if q_pub.strip():
            word = q_pub.strip()
            pat = re.compile(rf"({re.escape(word)})", re.IGNORECASE)
            res = df_pub[df_pub['Gap'].apply(lambda x: bool(pat.search(x)))] if not df_pub.empty else pd.DataFrame()
            if not res.empty:
                st.success(f"🔍 Publististik korpus bo'yicha jami {len(res)} ta mos gap aniqlandi.")
                for _, r in res.iterrows():
                    highlighted_sentence = pat.sub(r'<span class="highlight">\g<1></span>', r['Gap'])
                    st.markdown(f'<div class="sentence-container">{highlighted_sentence}</div>', unsafe_allow_html=True)
                    with st.expander("Metama'lumotlar"): 
                        st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
            else: st.warning("Publististik korpus matnlari ichida ushbu so'z topilmadi.")

    with tab_p2:
        st.subheader("📋 Talaba korpus asosida quyidagi jihatlarni tahlil qiladi:")
        top_words, top_bigrams = analyze_text_statistics(df_pub)
        
        col_an1, col_an2 = st.columns(2)
        with col_an1:
            st.info("### 1. Nutq strategiyalari (Sifat tahlili)\n\n* Talaba matnlarni o'qib persuaziv strategiyalarni ajratadi.\n* Baholovchi va emotsional ifodalarni KWIC orqali qidiradi.\n* Argumentatsiya usullarini aniqlaydi.")
            st.warning("### 🚀 Avtomatik Hisoblangan Kalit va Eng Ko'p Ishlatilgan So'zlar\n\nMatnlar tarkibidan avtomatik ravishda ajratib olingan chastotali lug'at:")
            if top_words:
                df_words = pd.DataFrame(top_words, columns=["So'z (Birlik)", "Uchrashish soni"])
                st.dataframe(df_words, use_container_width=True)
            else:
                st.write("Matnlar yuklanmagan.")
                
        with col_an2:
            st.info("### 2. Til birliklari (Miqdoriy tahlil)\n\nQuyidagi jadvallarda publististik bazadagi eng faol birliklar avtomatik ko'rsatilgan:")
            st.success("### 🔗 Avtomatik aniqlangan Kollokatsiyalar va Iboralar\n\nBirga eng ko'p kelgan so'z juftliklari:")
            if top_bigrams:
                df_bigrams = pd.DataFrame(top_bigrams, columns=["Kollokatsiya (Birikma)", "Chastotasi"])
                st.dataframe(df_bigrams, use_container_width=True)
            else:
                st.write("Birikmalar topilmadi.")

    with tab_p3:
        st.subheader("🏛️ Matnlardagi g'oyaviy va sotsiolingvistik tahlil yo'nalishlari:")
        st.warning("### 4-bosqich. Ideologik va ijtimoiy ma'nolarni aniqlash\n\n1. **Muallifning pozitsiyasi:** Matndagi sub'ektiv munosabat.\n2. **Ijtimoiy yoki siyosiy qarashlar:** Davr va mafkura tahlili.\n3. **Auditoriyaga ta'sir qilish usullari:** Manipulyativ va ritorik vositalar.\n4. **Diskursdagi asosiy g'oya:** Konseptual yadro.")