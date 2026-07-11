"""
Guepedia AI Editor Assistant
----------------------------
Aplikasi internal berbasis web (Streamlit) untuk membantu tim editor:
  1. Mengunggah naskah (.docx/.pdf/.txt)
  2. Mengekstrak teks & menghitung jumlah kata
  3. Mengklasifikasi genre & menghasilkan 3 opsi blurb via Groq API
  4. Meninjau, memilih/mengedit, dan menyimpan blurb terpilih ke SQLite

Jalankan dengan:  streamlit run app.py
"""
import os

import streamlit as st
from dotenv import load_dotenv

from db import init_db, save_analysis, update_chosen_blurb
from groq_client import GENRE_LIST, GroqRequestError, analyze_naskah
from utils import count_words, extract_text, guess_title

load_dotenv()

st.set_page_config(
    page_title="Guepedia AI Editor Assistant",
    page_icon="✨",
    layout="wide",
)

init_db()

# ---------------------------------------------------------------------------
# Styling — nuansa gelap/space sesuai identitas presentasi
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
.stApp {
    background: radial-gradient(
        ellipse at 15% 0%,
        #1f5f4a 0%,
        #12392e 40%,
        #081814 100%
    );
    color: #eefcf6;
}

/* Hilangkan sidebar */
section[data-testid="stSidebar"] {
    display: none;
}

/* Header */
.aig-eyebrow{
    font-size:12px;
    letter-spacing:3px;
    color:#7ef0b4;
    font-weight:700;
    text-transform:uppercase;
    margin-bottom:2px;
}

.aig-title{
    font-size:28px;
    font-weight:800;
    color:white;
    margin:0 0 6px 0;
}

.aig-sub{
    color:#b9e7d1;
    font-size:14.5px;
    margin-bottom:24px;
}

/* Card */
.aig-card{
    background:rgba(255,255,255,0.05);
    border:1px solid rgba(126,240,180,.18);
    border-radius:16px;
    padding:22px;
    backdrop-filter:blur(8px);
}

.aig-label{
    font-size:11px;
    color:#8fd7b5;
    text-transform:uppercase;
    letter-spacing:1px;
}

.aig-value{
    font-size:15px;
    margin-bottom:14px;
    color:white;
}

.aig-footnote{
    color:#8cb9a8;
    font-size:11.5px;
    text-align:center;
    margin-top:40px;
    line-height:1.6;
}

/* Text Area */
div[data-testid="stTextArea"] textarea{
    background:rgba(255,255,255,.05);
    color:white;
    border:1px solid rgba(126,240,180,.25);
    border-radius:12px;
}

/* Selectbox */
div[data-baseweb="select"]{
    background:rgba(255,255,255,.05);
    border-radius:10px;
}

/* Button */
.stButton > button{
    border-radius:10px;
    font-weight:600;
    transition:.25s;
}

.stButton > button:hover{
    transform:translateY(-2px);
}

/* Primary Button */
.stButton > button[kind="primary"]{
    background:linear-gradient(135deg,#53d88d,#1ea96c);
    color:white;
    border:none;
}

/* Success Alert */
div[data-testid="stAlert"]{
    border-radius:12px;
}
</style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
defaults = {
    "page": "dashboard",       # dashboard | result
    "raw_text": "",
    "file_name": "",
    "pending_genre_hint": "",
    "result": None,            # dict: genre, title, word_count, blurbs, record_id
    "active_blurb": 0,
    "error_msg": "",
}
for k, v in defaults.items():
    st.session_state.setdefault(k, v)


def run_analysis():
    st.session_state.error_msg = ""
    text = st.session_state.raw_text
    if count_words(text) < 5:
        st.session_state.error_msg = "Naskah terlalu pendek untuk dianalisis. Unggah file atau tempel teks naskah terlebih dahulu."
        return
    with st.spinner("Menganalisis naskah & menyusun blurb via Groq…"):
        try:
            analysis = analyze_naskah(text, genre_hint=st.session_state.pending_genre_hint)
        except GroqRequestError as e:
            st.session_state.error_msg = str(e)
            return

    title = guess_title(text)
    word_count = count_words(text)
    record_id = save_analysis(
        st.session_state.file_name or "(tempel manual)",
        title,
        word_count,
        analysis["genre"],
        analysis["blurbs"],
    )
    st.session_state.result = {
        "record_id": record_id,
        "title": title,
        "word_count": word_count,
        "genre": analysis["genre"],
        "blurbs": analysis["blurbs"],
    }
    st.session_state.active_blurb = 0
    st.session_state.page = "result"


def back_to_dashboard():
    st.session_state.page = "dashboard"
    st.session_state.error_msg = ""


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
col_logo, col_title = st.columns([0.08, 0.92])
with col_logo:
    st.image("asset/logo.jpg", width=140)
with col_title:
    st.markdown("<div class='aig-eyebrow'>Guepedia</div>", unsafe_allow_html=True)
    st.markdown("<div class='aig-title'>AI Editor Assistant</div>", unsafe_allow_html=True)

if not os.getenv("GROQ_API_KEY"):
    st.warning(
        "GROQ_API_KEY belum ditemukan di environment. Salin `.env.example` menjadi `.env`, "
        "isi API key dari https://console.groq.com/keys, lalu jalankan ulang `streamlit run app.py`.",
        icon="⚠️",
    )

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
if st.session_state.page == "dashboard":
    st.markdown("### Analisis Naskah Terbaru")
    st.markdown(
        "<div class='aig-sub'>Unggah dokumen naskah masuk untuk kalkulasi kata, klasifikasi genre, "
        "dan pembuatan blurb otomatis.</div>",
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Seret & lepas file naskah (.docx / .pdf / .txt)",
        type=["docx", "pdf", "txt"],
    )
    if uploaded is not None:
        try:
            st.session_state.raw_text = extract_text(uploaded)
            st.session_state.file_name = uploaded.name
            st.success(f"{uploaded.name} terbaca ({count_words(st.session_state.raw_text):,} kata)".replace(",", "."))
        except Exception as e:  # noqa: BLE001
            st.error(f"Gagal membaca file: {e}")

    st.markdown("<div style='text-align:center;color:#5f7099;font-size:12px;margin:10px 0;'>— atau tempel teks naskah secara manual —</div>", unsafe_allow_html=True)

    st.session_state.raw_text = st.text_area(
        "Isi naskah",
        value=st.session_state.raw_text,
        height=180,
        label_visibility="collapsed",
        placeholder="Tempel isi naskah di sini…",
    )

    # c1, c2 = st.columns([0.7, 0.3])
    # with c1:
    st.session_state.pending_genre_hint = st.selectbox(
            "Kategori genre (opsional)",
            options=[""] + GENRE_LIST,
            format_func=lambda g: "Pilih Genre (auto-deteksi)" if g == "" else g,
        )
    # with c2:
    st.write("")
    st.button("✨ Mulai Analisis", type="primary", use_container_width=True, on_click=run_analysis)

    if st.session_state.error_msg:
        st.error(st.session_state.error_msg)
        if "Groq" in st.session_state.error_msg or "GROQ" in st.session_state.error_msg:
            st.button("🔁 Klik Retry", on_click=run_analysis)

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
elif st.session_state.page == "result" and st.session_state.result:
    st.button("← Kembali ke Dashboard", on_click=back_to_dashboard)
    res = st.session_state.result

    col_a, col_b = st.columns([1, 1.3])

    with col_a:
        # st.markdown("<div class='aig-card'>", unsafe_allow_html=True)
        st.markdown("**Summary**")
        st.markdown(f"<div class='aig-label'>Judul</div><div class='aig-value'>{res['title']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='aig-label'>Jumlah Kata</div><div class='aig-value'>{res['word_count']:,}</div>".replace(",", "."), unsafe_allow_html=True)
        st.markdown(f"<div class='aig-label'>Klasifikasi Genre</div><div class='aig-value'>{res['genre']}</div>", unsafe_allow_html=True)
        # st.markdown("</div>", unsafe_allow_html=True)

    with col_b:
        # st.markdown("<div class='aig-card'>", unsafe_allow_html=True)
        st.markdown("**Generated Blurb**")
        tabs = st.tabs([f"Opsi {i+1}" for i in range(len(res["blurbs"]))])
        for i, tab in enumerate(tabs):
              with tab:
                edited = st.text_area(
                    f"blurb_{i}",
                    value=res["blurbs"][i],
                    height=140,
                    label_visibility="collapsed",
                    key=f"blurb_edit_{res['record_id']}_{i}",
                )

                st.markdown("**Copy Teks**")
                st.code(edited, language=None)

                if st.button(
                    "💾 Simpan Blurb Ini",
                    key=f"save_{i}",
                    type="primary",
                    use_container_width=True,
                ):
                    res["blurbs"][i] = edited
                    update_chosen_blurb(res["record_id"], i)
                    st.session_state.active_blurb = i
                    st.success("Blurb tersimpan ke database.")
                
        #     with tab:
        #         edited = st.text_area(
        #             f"blurb_{i}", value=res["blurbs"][i], height=140, label_visibility="collapsed", key=f"blurb_edit_{res['record_id']}_{i}"
        #         )
        #         col_save, col_copy = st.columns([0.5, 0.5])
        #         with col_copy:
        #             st.code(edited, language=None)
        #         with col_save:
        #             if st.button("💾 Simpan Blurb Ini", key=f"save_{i}"):
        #                 res["blurbs"][i] = edited
        #                 update_chosen_blurb(res["record_id"], i)
        #                 st.session_state.active_blurb = i
        #                 st.success("Blurb tersimpan ke database.")
        #         # with col_copy:
        #         #     st.code(edited, language=None)
        # # st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    "<div class='aig-footnote'>Guepedia Editor Assistant 2026</div>",
    unsafe_allow_html=True,
)
