import os
import re
from collections import Counter
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="O'ZBEK TILI KORPUSI", layout="wide")

st.markdown("""
<style>
    .highlight { background-color: #FDE047 !important; color: #000000 !important; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
    .sentence-container { background-color: #FFFFFF; padding: 15px 20px; border-radius: 6px; border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 10px; font-size: 16px; color: #1E293B; }
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

UMUMIY_FOLDER = "data/umumiy" if os.path.exists("data/umumiy") else "umumiy"
PUBLISTISTIKA_FOLDER = "data/publististika" if os.path.exists("data/publististika") else "publististika"

st.sidebar.markdown("### 🧭 NAVIGATION")
page = st.sidebar.radio("Tanlang:", ["🏠 Bosh sahifa", "📂 Umumiy korpus", "🌐 Parallel korpus", "✍️ Publististik korpus"])

if page == "🏠 Bosh sahifa":
    st.title("O'ZBEK TILI KORPUSI")
    df_u = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    df_p = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    c1, c2, c3 = st.columns(3)
    c1.success(f"📂 Umumiy korpus\n\n{len(df_u)} ta gap")
    c2.info("🌐 Parallel korpus\n\nO'zbek-Turk tili")
    c3.warning(f"✍️ Publististik korpus\n\n{len(df_p)} ta gap")

elif page == "📂 Umumiy korpus":
    df = load_korpus_baza(UMUMIY_FOLDER, 24, "text_", "tag_", "txt", "docx")
    st.title("📂 Umumiy korpus")
    q = st.text_input("So'z kiriting:")
    if q.strip():
        pat = re.compile(rf"({re.escape(q.strip())}[b-df-hj-np-rt-vxz'aouei‘g‘shch]*)", re.IGNORECASE)
        res = df[df['Gap'].apply(lambda x: bool(pat.search(x)))] if not df.empty else pd.DataFrame()
        if not res.empty:
            st.success(f"{len(res)} ta gap topildi.")
            for _, r in res.iterrows():
                st.markdown(f'<div class="sentence-container">{pat.sub(r"<span class=\"highlight\">\g<1></span>", r["Gap"])}</div>', unsafe_allow_html=True)
                with st.expander("Metama'lumotlar"): st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
        else: st.warning("Topilmadi.")

elif page == "🌐 Parallel korpus":
    st.title("🌐 O'zbek-Turk Parallel Korpusi")
    st.components.v1.iframe("https://uzbek-turk-parallel-korpusi-cnzm5cmc3tkccaysyxai5s.streamlit.app/?embed=true", height=800)

elif page == "✍️ Publististik korpus":
    df_pub = load_korpus_baza(PUBLISTISTIKA_FOLDER, 21, "pub.", "teg.", "txt", "docx")
    st.title("✍️ Publististik matnlar korpusi")
    t1, t2, t3 = st.tabs(["🔍 KWIC Qidiruv", "📊 Diskurs tahlil", "🏛️ Mafkuraviy tahlil"])
    
    with t1:
        q_p = st.text_input("Publististikadan so'z kiriting:")
        if q_p.strip():
            pat = re.compile(rf"({re.escape(q_p.strip())}[b-df-hj-np-rt-vxz'aouei‘g‘shch]*)", re.IGNORECASE)
            res = df_pub[df_pub['Gap'].apply(lambda x: bool(pat.search(x)))] if not df_pub.empty else pd.DataFrame()
            if not res.empty:
                st.success(f"{len(res)} ta gap topildi.")
                for _, r in res.iterrows():
                    st.markdown(f'<div class="sentence-container">{pat.sub(r"<span class=\"highlight\">\g<1></span>", r["Gap"])}</div>', unsafe_allow_html=True)
                    with st.expander("Metama'lumotlar"): st.write({k: v for k, v in r.to_dict().items() if k not in ["Gap", "Fayl"]})
            else: st.warning("Topilmadi.")
            
    with t2:
        words, bigrams = analyze_text_statistics(df_pub)
        col1, col2 = st.columns(2)
        col1.write("### 📈 Eng ko'p ishlatilgan so'zlar:")
        col1.dataframe(pd.DataFrame(words, columns=["So'z", "Soni"]), use_container_width=True)
        col2.write("### 🔗 Avtomatik kollokatsiyalar:")
        col2.dataframe(pd.DataFrame(bigrams, columns=["Birikma", "Chastota"]), use_container_width=True)
        
    with t3:
        st.info("### 🏛️ 4-bosqich. Ideologik va ijtimoiy ma'nolarni aniqlash\n\n1. **Muallifning pozitsiyasi:** Matndagi sub'ektiv munosabat.\n2. **Ijtimoiy yoki siyosiy qarashlar:** Davr va mafkura tahlili.\n3. **Auditoriyaga ta'sir qilish usullari:** Ritorik vositalar.\n4. **Diskursdagi asosiy g'oya:** Konseptual yadro.")
