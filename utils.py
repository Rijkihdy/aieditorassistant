"""
Fungsi bantu: ekstraksi naskah (docx/pdf/txt) & kalkulasi jumlah kata.
"""
import io


def extract_text(uploaded_file):
    """
    uploaded_file: objek dari st.file_uploader (punya .name dan bisa dibaca bytes-nya).
    Mengembalikan teks polos hasil ekstraksi.
    """
    name = uploaded_file.name.lower()
    raw_bytes = uploaded_file.read()

    if name.endswith(".txt"):
        return raw_bytes.decode("utf-8", errors="ignore")

    if name.endswith(".docx"):
        from docx import Document

        doc = Document(io.BytesIO(raw_bytes))
        return "\n".join(p.text for p in doc.paragraphs)

    if name.endswith(".pdf"):
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(raw_bytes))
        text_parts = [page.extract_text() or "" for page in reader.pages]
        return "\n".join(text_parts)

    raise ValueError(f"Format file '{name}' belum didukung. Gunakan .docx, .pdf, atau .txt.")


def count_words(text: str) -> int:
    text = text.strip()
    if not text:
        return 0
    return len(text.split())


def guess_title(text: str, fallback: str = "Naskah Tanpa Judul") -> str:
    for line in text.splitlines():
        line = line.strip()
        if line:
            return line[:80]
    return fallback
