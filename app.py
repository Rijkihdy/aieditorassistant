"""
Guepedia AI Editor Assistant
----------------------------
Aplikasi internal berbasis web (Streamlit) untuk membantu tim editor:
  1. Mengunggah naskah (.docx/.pdf/.txt)
  2. Mengekstrak teks & menghitung jumlah kata
  3. Mengklasifikasi genre & menghasilkan 3 opsi blurb via Groq API
  4. Meninjau, memilih/mengedit, dan menyimpan blurb terpilih ke SQLite
  5. Menyesuaikan format dinamis & generate dokumen final (.docx)

Jalankan dengan:  streamlit run app.py
"""
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path
from cover_ai.prompt_builder import build_prompt, hitung_spine
from cover_ai.gemini_client import generate_cover

import streamlit as st
from dotenv import load_dotenv

from db import init_db, save_analysis, update_chosen_blurb
from docx_formatter import normalize_docx_layout
from groq_client import GENRE_LIST, GroqRequestError, analyze_naskah, generate_synopsis
from mailmerge import MailMergeError, generate_book_docx
from naskah_parser import parse_naskah
from utils import count_words, extract_text, guess_title

DEFAULT_TEMPLATE_PATH = Path(__file__).parent / "templates" / "template_buku.docx"

load_dotenv()

st.set_page_config(
    page_title="Guepedia AI Editor Assistant",
    page_icon="✨",
    layout="wide",
)

init_db()

# ---------------------------------------------------------------------------
# Styling
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

section[data-testid="stSidebar"] {
    display: none;
}

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

.aig-card{
    background:rgba(255,255,255,0.05);
    border:1px solid rgba(126,240,180,.18);
    border-radius:16px;
    padding:22px;
    backdrop-filter:blur(8px);
}

.aig-panel{
    background:rgba(255,255,255,0.04);
    border:1px solid rgba(126,240,180,.14);
    border-radius:18px;
    padding:18px 20px 22px 20px;
    margin-bottom:18px;
}

.aig-step{
    background:rgba(255,255,255,0.04);
    border:1px solid rgba(126,240,180,.14);
    border-radius:14px;
    padding:14px;
    min-height:120px;
}

.aig-step-number{
    display:inline-block;
    width:28px;
    height:28px;
    line-height:28px;
    text-align:center;
    border-radius:999px;
    background:linear-gradient(135deg,#53d88d,#1ea96c);
    color:white;
    font-weight:700;
    margin-bottom:8px;
}

.aig-step-title{
    font-size:13px;
    font-weight:700;
    color:white;
    margin-bottom:4px;
}

.aig-step-desc{
    font-size:12px;
    color:#bfdccf;
    line-height:1.5;
}

.aig-pill{
    display:inline-block;
    padding:4px 8px;
    border-radius:999px;
    background:rgba(126,240,180,.16);
    color:#bff4d0;
    font-size:11px;
    font-weight:700;
    letter-spacing:.5px;
    text-transform:uppercase;
    margin-bottom:8px;
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

div[data-testid="stTextArea"] textarea{
    background:rgba(255,255,255,.05);
    color:white;
    border:1px solid rgba(126,240,180,.25);
    border-radius:12px;
}

div[data-baseweb="select"]{
    background:rgba(255,255,255,.05);
    border-radius:10px;
}

.stButton > button{
    border-radius:10px;
    font-weight:600;
    transition:.25s;
}

.stButton > button:hover{
    transform:translateY(-2px);
}

.stButton > button[kind="primary"]{
    background:linear-gradient(135deg,#53d88d,#1ea96c);
    color:white;
    border:none;
}

div[data-testid="stAlert"]{
    border-radius:12px;
}
</style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# State Defaults
# ---------------------------------------------------------------------------
defaults = {
    "page": "dashboard",
    "cover_prompt": "",
    "raw_text": "",
    "file_name": "",
    "pending_genre_hint": "",
    "result": None,
    "active_blurb": 0,
    "error_msg": "",
    "formatted_docx_bytes": None,
    "editor_sections": [],
    "editor_source_text": "",
    "fmt_margin_top": 2.0,
    "fmt_margin_bottom": 2.0,
    "fmt_margin_left": 2.5,
    "fmt_margin_right": 2.5,
    "fmt_font_name": "Times New Roman",
    "fmt_font_size": 12,
    "fmt_line_spacing": 1.15,
    "fmt_alignment": "justify",
    "fmt_header_text": "NAMA DOKUMEN - DIFORMAT OTOMATIS",
    "fmt_header_font_name": "Arial",
    "fmt_header_font_size": 9,
    "mm_nama_penulis": "",
    "mm_isbn": "",
    "mm_qrcbn": "",
    "mm_judul_naskah": "",
    "mm_tahun_cetak": "Mei 2025",
    "mm_kata_pengantar": "",
    "fmt_heading_alignment": "left",
    "sinopsis_mode": "Tulis/Upload Manual",
    "sinopsis_manual_text": "",
    "sinopsis_ai_text": "",
}
for k, v in defaults.items():
    st.session_state.setdefault(k, v)


def convert_docx_bytes(uploaded_file, config: dict) -> bytes:
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
        tmp_in.write(uploaded_file.read())
        tmp_input_path = tmp_in.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
        tmp_output_path = tmp_out.name

    normalize_docx_layout(tmp_input_path, tmp_output_path, config)

    with open(tmp_output_path, "rb") as f:
        formatted_bytes = f.read()
    
    st.session_state.formatted_docx_bytes = formatted_bytes
    return formatted_bytes


def extract_text_from_docx_bytes(docx_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    return "\n".join(p.text for p in doc.paragraphs)


def build_editor_sections(raw_text: str, fallback_title: str | None = None) -> list[dict[str, str]]:
    if not raw_text or not raw_text.strip():
        return []

    kata_pengantar, sections = parse_naskah(raw_text, fallback_title=fallback_title or "")

    if kata_pengantar and not st.session_state.get("mm_kata_pengantar"):
        st.session_state["mm_kata_pengantar"] = kata_pengantar

    return [{"title": title.strip(), "body": "\n".join(body).strip()} for title, body in sections]


def serialize_editor_sections(sections: list[dict[str, str]]) -> str:
    lines: list[str] = []
    for section in sections:
        title = (section.get("title") or "").strip()
        body = (section.get("body") or "").strip()
        if title:
            lines.append(title)
        if body:
            lines.extend(line.strip() for line in body.splitlines() if line.strip())
    return "\n".join(lines)


def generate_book_bytes(template_file, meta: dict, chapter_title: str, kata_pengantar_text: str, naskah_text: str, format_config: dict | None = None, source_docx_bytes: bytes | None = None, qrcbn: str = "", sinopsis_text: str = "") -> bytes:
    if template_file is not None:
        if hasattr(template_file, "seek"):
            template_file.seek(0)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_tpl:
            tmp_tpl.write(template_file.read())
            template_path = tmp_tpl.name
    else:
        template_path = str(DEFAULT_TEMPLATE_PATH)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
        output_path = tmp_out.name

    generate_book_docx(
        template_path=template_path,
        output_path=output_path,
        nama_penulis=meta["nama_penulis"],
        judul_naskah=meta["judul_naskah"],
        isbn=meta["isbn"],
        tahun_cetak=meta["tahun_cetak"],
        kata_pengantar_text=kata_pengantar_text,
        chapter_title=chapter_title,
        naskah_text=naskah_text,
        format_config=format_config,
        source_docx_bytes=source_docx_bytes,
        qrcbn=qrcbn,
        sinopsis_text=sinopsis_text,
    )

    with open(output_path, "rb") as f:
        return f.read()


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

if not os.getenv("GEMINI_API_KEY"):
    st.warning(
        "GEMINI_API_KEY belum ditemukan di environment. Fitur AI Cover Generator butuh ini. "
        "Daftar gratis di https://aistudio.google.com/apikey, isi ke file `.env` "
        "(GEMINI_API_KEY=xxxxx), lalu jalankan ulang `streamlit run app.py`.",
        icon="⚠️",
    )

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
if st.session_state.page == "dashboard":
    st.markdown(
        "<div class='aig-panel'>"
        "<div class='aig-pill'>Workflow</div>"
        "<div class='aig-sub'>Pilih naskah, jalankan analisis, lalu review hasil blurb dalam satu alur yang lebih terarah.</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    col_step1, col_step2, col_step3 = st.columns(3)
    with col_step1:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>1</div><div class='aig-step-title'>Input Naskah</div><div class='aig-step-desc'>Unggah file atau tempel teks langsung ke area input.</div></div>",
            unsafe_allow_html=True,
        )
    with col_step2:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>2</div><div class='aig-step-title'>Analisis Otomatis</div><div class='aig-step-desc'>Sistem menghitung kata, menilai genre, dan menghasilkan opsi blurb.</div></div>",
            unsafe_allow_html=True,
        )
    with col_step3:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>3</div><div class='aig-step-title'>Review & Export</div><div class='aig-step-desc'>Pilih blurb terbaik, simpan hasil, lalu ekspor ke dokumen final.</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    col_main, col_side = st.columns([1.2, 0.8], gap="large")

    with col_main:
        st.markdown("### 1. Input & Analisis")
        st.markdown(
            "<div class='aig-sub'>Unggah dokumen naskah atau tempel teks secara manual, lalu jalankan analisis.</div>",
            unsafe_allow_html=True,
        )

        uploaded = st.file_uploader(
            "Seret & lepas file naskah (.docx / .pdf / .txt)",
            type=["docx", "pdf", "txt"],
            key="main_naskah_uploader",
        )
        if uploaded is not None:
            try:
                st.session_state.raw_text = extract_text(uploaded)
                st.session_state.file_name = uploaded.name
                st.success(f"{uploaded.name} terbaca ({count_words(st.session_state.raw_text):,} kata)".replace(",", "."))
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")

        st.markdown("<div style='text-align:center;color:#5f7099;font-size:12px;margin:10px 0;'>— atau tempel teks naskah secara manual —</div>", unsafe_allow_html=True)

        st.session_state.raw_text = st.text_area(
            "Isi naskah",
            value=st.session_state.raw_text,
            height=180,
            label_visibility="collapsed",
            placeholder="Tempel isi naskah di sini…",
        )

        if st.session_state.raw_text.strip():
            st.info(f"Naskah siap diproses — sekitar {count_words(st.session_state.raw_text):,} kata".replace(",", "."))
        else:
            st.caption("Belum ada naskah yang dimasukkan. Silakan unggah file atau tempel teks.")

        if st.session_state.raw_text.strip():
            fallback_title = st.session_state.result["title"] if st.session_state.result else guess_title(st.session_state.raw_text)
            if st.session_state.get("editor_source_text") != st.session_state.raw_text:
                st.session_state.editor_sections = build_editor_sections(st.session_state.raw_text, fallback_title=fallback_title)
                st.session_state.editor_source_text = st.session_state.raw_text

            if st.session_state.editor_sections:
                st.markdown("### 3. Editor otomatis naskah")
                st.caption("Sistem memecah naskah menjadi bagian yang bisa diedit per bab atau sub-bab secara otomatis.")
                for idx, section in enumerate(st.session_state.editor_sections):
                    with st.expander(f"{section['title'] or f'Bagian {idx + 1}'}", expanded=idx == 0):
                        title_value = st.text_input(
                            "Judul bagian",
                            value=section.get("title", ""),
                            key=f"editor_title_{idx}",
                        )
                        body_value = st.text_area(
                            "Isi bagian",
                            value=section.get("body", ""),
                            height=140,
                            key=f"editor_body_{idx}",
                        )
                        st.session_state.editor_sections[idx]["title"] = title_value
                        st.session_state.editor_sections[idx]["body"] = body_value

        st.session_state.pending_genre_hint = st.selectbox(
            "Kategori genre (opsional)",
            options=[""] + GENRE_LIST,
            format_func=lambda g: "Pilih Genre (auto-deteksi)" if g == "" else g,
        )

        st.write("")
        st.button("✨ Mulai Analisis", type="primary", use_container_width=True, on_click=run_analysis)

        if st.session_state.error_msg:
            st.error(st.session_state.error_msg)
            if "Groq" in st.session_state.error_msg or "GROQ" in st.session_state.error_msg:
                st.button("🔁 Klik Retry", on_click=run_analysis)
        st.divider()

        st.subheader("🎨 AI Cover Generator")
        col1, col2 = st.columns(2)
        with col1:

            jenis_buku = st.radio(
                "Jenis Buku",
                ["nonfiksi", "fiksi"],
                horizontal=True,
            )

            judul = st.text_input(
                "Judul Buku",
                value=st.session_state.get("mm_judul_naskah", ""),
            )

            subjudul = st.text_input(
                "Sub Judul",
                placeholder="Kosongkan jika tidak ada"
            )

            penulis = st.text_input(
                "Nama Penulis",
                value=st.session_state.get("mm_nama_penulis", ""),
            )

            kategori = st.text_input(
                "Kategori Buku",
                value=st.session_state.get("pending_genre_hint", "")
            )

            genre = st.text_input(
                "Genre Buku",
                value=st.session_state.get("pending_genre_hint", "")
            )

            jumlah_halaman = st.number_input(
                "Jumlah Halaman",
                min_value=1,
                value=100,
            )
        with col2:
            spine = hitung_spine(jumlah_halaman)
            st.metric(
                "Lebar Spine",
                f"{spine:.2f} mm"
            )

        sinopsis_final = (
            st.session_state.get("sinopsis_manual_text", "")
            or st.session_state.get("sinopsis_ai_text", "")
        )

        if st.button("🎨 Generate Cover"):

            field_wajib = {
                "Judul Buku": judul,
                "Nama Penulis": penulis,
                "Sinopsis": sinopsis_final,
            }
            kosong = [nama for nama, isi in field_wajib.items() if not isi.strip()]

            if kosong:
                st.warning(
                    "Sesuai brief, field berikut wajib diisi dulu sebelum generate cover: "
                    + ", ".join(kosong)
                    + ". Sinopsis bisa diisi manual atau di-generate AI di panel "
                    "'📝 Sinopsis' pada sidebar kanan."
                )
            else:
                try:
                    prompt = build_prompt(
                        jenis_buku=jenis_buku,
                        judul=judul,
                        subjudul=subjudul,
                        penulis=penulis,
                        kategori=kategori,
                        genre=genre,
                        sinopsis=sinopsis_final,
                        jumlah_halaman=jumlah_halaman,
                    )
                except Exception as e:
                    st.error(f"Gagal membangun prompt: {e}")
                    prompt = None

                if prompt is not None:
                    # Simpan prompt ke session state
                    st.session_state.cover_prompt = prompt
                    st.success("✅ Prompt berhasil dibuat.")

                    try:
                        image = generate_cover(prompt)
                    except Exception as e:
                        st.error(f"Gagal generate gambar cover: {e}")
                        image = None

                    if image is not None:
                        st.image(
                            image,
                            caption="Preview Cover",
                            use_container_width=True
                        )

                    with st.expander("🔍 Lihat Prompt"):
                        st.text_area(
                            "Prompt Final",
                            value=st.session_state.cover_prompt,
                            height=450,
                            disabled=True,
                        )

    with col_side:
        st.markdown("### 2. Format & Ekspor")
        st.markdown(
            "<div class='aig-sub'>Fitur pendukung ini dipisahkan agar alur utama tetap fokus dan tidak terlalu padat.</div>",
            unsafe_allow_html=True,
        )

        with st.expander("📝 Sinopsis", expanded=False):
            st.caption("Pilih sumber sinopsis untuk halaman depan buku.")
            sinopsis_mode = st.radio(
                "Sumber Sinopsis",
                options=["Tulis/Upload Manual", "Generate dengan AI"],
                key="sinopsis_mode",
                horizontal=True,
            )

            if sinopsis_mode == "Tulis/Upload Manual":
                st.text_area(
                    "Isi Sinopsis",
                    height=150,
                    placeholder="Tempel atau tulis sinopsis di sini…",
                    key="sinopsis_manual_text",
                )
            else:
                st.caption("AI meringkas SEMUA bab (hasil Editor Otomatis Naskah di atas) jadi satu sinopsis 50–100 kalimat.")
                if st.button("🤖 Generate Sinopsis dengan AI", use_container_width=True):
                    if not st.session_state.editor_sections:
                        st.warning("Unggah/tempel naskah dulu supaya Editor Otomatis Naskah bisa memecahnya per bab.")
                    else:
                        with st.spinner("Meringkas seluruh bab menjadi sinopsis…"):
                            try:
                                sections_for_ai = [
                                    (s.get("title", ""), s.get("body", "").splitlines())
                                    for s in st.session_state.editor_sections
                                ]
                                judul_for_ai = st.session_state.get("mm_judul_naskah") or (
                                    st.session_state.result["title"] if st.session_state.result else ""
                                )
                                st.session_state.sinopsis_ai_text = generate_synopsis(sections_for_ai, judul_for_ai)
                                st.success("Sinopsis berhasil dibuat. Review/edit hasilnya di bawah kalau perlu.")
                            except GroqRequestError as e:
                                st.error(str(e))
                st.text_area(
                    "Hasil Sinopsis AI (bisa diedit)",
                    height=150,
                    key="sinopsis_ai_text",
                )

        with st.expander("📄 Formatter & Siapkan Dokumen", expanded=True):
            st.markdown(
                "<div class='aig-sub'>Satu alur: atur format dinamis, isi metadata halaman depan, lalu langsung siapkan dokumen final.</div>",
                unsafe_allow_html=True,
            )
            with st.form("formatter_form"):
                st.markdown("#### 📏 Layout & Formatter Dinamis")
                cfg_margin_top = st.number_input("Margin Atas (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_top")
                cfg_margin_bottom = st.number_input("Margin Bawah (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_bottom")
                cfg_margin_left = st.number_input("Margin Kiri (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_left")
                cfg_margin_right = st.number_input("Margin Kanan (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_right")
                cfg_font_name = st.text_input("Nama Font Utama", key="fmt_font_name")
                cfg_font_size = st.number_input("Ukuran Font Utama (pt)", min_value=8, max_value=20, step=1, key="fmt_font_size")
                cfg_line_spacing = st.number_input("Line Spacing", min_value=1.0, max_value=2.0, step=0.05, key="fmt_line_spacing")
                cfg_alignment = st.selectbox("Alignment Isi Naskah", options=["justify", "left", "center", "right"], key="fmt_alignment")
                cfg_heading_alignment = st.selectbox(
                    "Alignment Header Bab",
                    options=["left", "center", "right"],
                    key="fmt_heading_alignment",
                    format_func=lambda a: {"left": "Kiri", "center": "Tengah", "right": "Kanan"}[a],
                )
                cfg_header_text = st.text_input("Teks Header", key="fmt_header_text")
                cfg_header_font_name = st.text_input("Font Header", key="fmt_header_font_name")
                cfg_header_font_size = st.number_input("Ukuran Font Header (pt)", min_value=6, max_value=16, step=1, key="fmt_header_font_size")

                st.markdown("#### ✍️ Halaman Depan & Redaksi")
                mm_template_file = st.file_uploader(
                    "Template buku kustom (opsional)",
                    type=["docx"],
                    key="mm_template_uploader",
                )

                col_mm1, col_mm2 = st.columns(2)
                with col_mm1:
                    mm_nama_penulis = st.text_input("Nama Penulis", key="mm_nama_penulis")
                    mm_isbn = st.text_input("ISBN", placeholder="978-623-xxx-xx-x", key="mm_isbn")
                    mm_qrcbn = st.text_input("QRCBN (opsional)", placeholder="62-xxxx-xxxx-xxx", key="mm_qrcbn")
                with col_mm2:
                    default_judul = st.session_state.result["title"] if st.session_state.result else ""
                    if default_judul and not st.session_state.get("mm_judul_naskah"):
                        st.session_state["mm_judul_naskah"] = default_judul
                    mm_judul_naskah = st.text_input("Judul Naskah", key="mm_judul_naskah")
                    mm_tahun_cetak = st.text_input("Tahun/Bulan Cetak", key="mm_tahun_cetak")

                mm_kata_pengantar = st.text_area(
                    "Isi Kata Pengantar (Opsional)",
                    height=120,
                    placeholder="Setiap manusia punya cara yang berbeda dalam mencintai, merelakan, dan bertahan...",
                    key="mm_kata_pengantar"
                )

                submitted_format = st.form_submit_button("📄 Format & Siapkan Dokumen", type="primary", use_container_width=True)
                if submitted_format:
                    if count_words(st.session_state.raw_text) < 5:
                        st.warning("Unggah file naskah atau tempel teks naskah terlebih dahulu di atas.")
                    elif not mm_nama_penulis or not mm_judul_naskah:
                        st.warning("Nama Penulis dan Judul Naskah wajib diisi.")
                    else:
                        try:
                            chapter_title = (
                                st.session_state.editor_sections[0]["title"]
                                if st.session_state.editor_sections
                                else (st.session_state.result["title"] if st.session_state.result else guess_title(st.session_state.raw_text))
                            )

                            source_docx_bytes = None
                            if uploaded is not None and uploaded.name.lower().endswith(".docx"):
                                config = {
                                    "margin_top_cm": cfg_margin_top,
                                    "margin_bottom_cm": cfg_margin_bottom,
                                    "margin_left_cm": cfg_margin_left,
                                    "margin_right_cm": cfg_margin_right,
                                    "font_name": cfg_font_name,
                                    "font_size_pt": cfg_font_size,
                                    "line_spacing": cfg_line_spacing,
                                    "alignment": cfg_alignment,
                                    "header_text": cfg_header_text,
                                    "header_font_name": cfg_header_font_name,
                                    "header_font_size_pt": cfg_header_font_size,
                                }
                                source_docx_bytes = convert_docx_bytes(uploaded, config)
                                st.session_state["formatted_docx_bytes"] = source_docx_bytes
                                st.session_state["formatted_docx_name"] = f"{Path(uploaded.name).stem}_normalized.docx"
                                naskah_text = extract_text_from_docx_bytes(source_docx_bytes)
                            else:
                                naskah_text = serialize_editor_sections(st.session_state.editor_sections) if st.session_state.editor_sections else st.session_state.raw_text

                            format_config = {
                                "margin_top_cm": cfg_margin_top,
                                "margin_bottom_cm": cfg_margin_bottom,
                                "margin_left_cm": cfg_margin_left,
                                "margin_right_cm": cfg_margin_right,
                                "font_name": cfg_font_name,
                                "font_size_pt": cfg_font_size,
                                "line_spacing": cfg_line_spacing,
                                "alignment": cfg_alignment,
                                "heading_alignment": cfg_heading_alignment,
                                "header_text": cfg_header_text,
                                "header_font_name": cfg_header_font_name,
                                "header_font_size_pt": cfg_header_font_size,
                            }

                            final_sinopsis_text = (
                                st.session_state.get("sinopsis_manual_text", "")
                                if st.session_state.get("sinopsis_mode") == "Tulis/Upload Manual"
                                else st.session_state.get("sinopsis_ai_text", "")
                            )

                            book_bytes = generate_book_bytes(
                                mm_template_file,
                                meta={
                                    "nama_penulis": mm_nama_penulis,
                                    "judul_naskah": mm_judul_naskah,
                                    "isbn": mm_isbn,
                                    "tahun_cetak": mm_tahun_cetak,
                                },
                                chapter_title=chapter_title,
                                kata_pengantar_text=mm_kata_pengantar,
                                naskah_text=naskah_text,
                                format_config=format_config,
                                source_docx_bytes=source_docx_bytes or st.session_state.formatted_docx_bytes,
                                qrcbn=mm_qrcbn,
                                sinopsis_text=final_sinopsis_text,
                            )
                            st.session_state["generated_book_bytes"] = book_bytes
                            st.session_state["generated_book_name"] = f"{mm_judul_naskah.strip() or 'naskah'}_final.docx"
                            st.success("Dokumen final berhasil disiapkan. Silakan unduh hasilnya di bawah.")
                        except MailMergeError as e:
                            st.error(f"Template tidak sesuai: {e}")
                        except Exception as e:
                            st.error(f"Gagal menyiapkan dokumen: {e}")

            if "generated_book_bytes" in st.session_state and st.session_state.get("generated_book_bytes") is not None:
                st.download_button(
                    "⬇️ Unduh Dokumen Final",
                    data=st.session_state["generated_book_bytes"],
                    file_name=st.session_state.get("generated_book_name", "dokumen_final.docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
elif st.session_state.page == "result" and st.session_state.result:
    st.button("← Kembali ke Dashboard", on_click=back_to_dashboard)
    res = st.session_state.result

    st.markdown("### Hasil Analisis")
    st.markdown(
        "<div class='aig-sub'>Review tiap opsi blurb, pilih yang paling sesuai, lalu simpan hasilnya.</div>",
        unsafe_allow_html=True,
    )

    col_a, col_b = st.columns([0.9, 1.1], gap="large")

    with col_a:
        summary_html = (
            "<div class='aig-card'>"
            "<div class='aig-pill'>Summary</div>"
            f"<div class='aig-label'>Judul</div><div class='aig-value'>{res['title']}</div>"
            f"<div class='aig-label'>Jumlah Kata</div><div class='aig-value'>{res['word_count']:,}</div>"
            f"<div class='aig-label'>Klasifikasi Genre</div><div class='aig-value'>{res['genre']}</div>"
            "</div>"
        )
        summary_html = summary_html.replace(",", ".")
        st.markdown(summary_html, unsafe_allow_html=True)

    with col_b:
        st.markdown(
            "<div class='aig-card'>"
            "<div class='aig-pill'>Generated Blurb</div>"
            "</div>",
            unsafe_allow_html=True,
        )
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

                if st.session_state.active_blurb == i:
                    st.success("Blurb ini sudah tersimpan.")
                else:
                    st.caption("Belum disimpan ke database.")

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

st.markdown(
    "<div class='aig-footnote'>Guepedia Editor Assistant 2026</div>",
    unsafe_allow_html=True,
)