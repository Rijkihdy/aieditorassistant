import io
from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def normalize_docx_layout(input_docx_path: str, output_docx_path: str, config: dict) -> None:
    """
    Normalize a messy .docx layout and write a cleaned version to output_docx_path.

    Args:
        input_docx_path: Path to the source .docx file.
        output_docx_path: Path to save the normalized .docx file.
        config: Layout configuration dictionary, for example:
            {
                "margin_top_cm": 2.0,
                "margin_bottom_cm": 2.0,
                "margin_left_cm": 2.5,
                "margin_right_cm": 2.5,
                "font_name": "Times New Roman",
                "font_size_pt": 12,
                "line_spacing": 1.15,
                "alignment": "justify",
                "header_text": "Nama Dokumen - Normalisasi Otomatis",
                "header_font_name": "Arial",
                "header_font_size_pt": 9,
            }

    This function preserves bold/italic run-level formatting and keeps Heading
    styles using their default/relative sizes.
    """

    document = Document(input_docx_path)

    # 1. Terapkan margin global pada setiap section.
    _apply_margins_to_all_sections(document, config)

    # 2. Terapkan font, spasi, dan alignment pada semua paragraf dan tabel.
    _normalize_document_text(document, config)

    # 3. Hapus paragraf kosong yang hanya berisi garis baru atau whitespace.
    _clean_empty_paragraphs(document)

    # 4. Tambahkan header kustom di section pertama.
    _insert_custom_header(document, config)

    # 5. Inject nomor halaman dinamis menggunakan OXML di footer.
    _inject_dynamic_page_numbering(document)

    Path(output_docx_path).parent.mkdir(parents=True, exist_ok=True)
    document.save(output_docx_path)


def _apply_margins_to_all_sections(document: Document, config: dict) -> None:
    """Set margins for every section in the document."""
    margin_map = {
        "top": config.get("margin_top_cm", 2.0),
        "bottom": config.get("margin_bottom_cm", 2.0),
        "left": config.get("margin_left_cm", 2.5),
        "right": config.get("margin_right_cm", 2.5),
    }

    for section in document.sections:
        section.top_margin = Cm(margin_map["top"])
        section.bottom_margin = Cm(margin_map["bottom"])
        section.left_margin = Cm(margin_map["left"])
        section.right_margin = Cm(margin_map["right"])


def _normalize_document_text(document: Document, config: dict) -> None:
    """Apply font settings, line spacing, and paragraph alignment globally.

    PENGECUALIAN PENTING: paragraf di halaman COVER dan halaman IDENTITAS
    BUKU (semua yang ada SEBELUM heading "KATA PENGANTAR") sengaja TIDAK
    ditimpa alignment-nya. Halaman-halaman itu didesain center-aligned di
    template, dan kalau alignment isi naskah (biasanya "justify") diterapkan
    global ke SEMUA paragraf, halaman cover/identitas ikut berubah jadi rata
    kiri-kanan padahal seharusnya tetap center.
    """
    font_name = config.get("font_name")
    font_size_pt = config.get("font_size_pt")
    line_spacing = config.get("line_spacing")
    alignment = config.get("alignment")

    alignment_map = {
        "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
    }

    in_front_matter = True

    for paragraph in _iter_all_paragraphs(document):
        if _is_field_paragraph(paragraph):
            continue

        if in_front_matter and paragraph.text.strip().upper() == "KATA PENGANTAR":
            in_front_matter = False

        is_heading = _is_heading_style(paragraph)

        if line_spacing is not None and not in_front_matter:
            paragraph.paragraph_format.line_spacing = line_spacing

        if alignment and not in_front_matter:
            paragraph.alignment = alignment_map.get(alignment.lower(), paragraph.alignment)

        for run in paragraph.runs:
            if font_name:
                run.font.name = font_name
            if font_size_pt is not None and not is_heading:
                run.font.size = Pt(font_size_pt)

        # Pastikan tabel juga mempertahankan styling di dalam sel.
        if paragraph._p.getparent().tag == qn("w:tc"):
            continue


def _is_heading_style(paragraph) -> bool:
    """Return True when the paragraph uses a Heading style."""
    style_name = ""
    if paragraph.style is not None:
        style_name = paragraph.style.name or ""
    return style_name.startswith("Heading")


def _iter_all_paragraphs(document: Document):
    """Yield all paragraphs from the document body and all tables."""
    for paragraph in document.paragraphs:
        yield paragraph

    for table in document.tables:
        yield from _iter_table_paragraphs(table)


def _iter_table_paragraphs(table):
    """Recursively yield paragraphs inside a table."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                yield paragraph
            for nested_table in cell.tables:
                yield from _iter_table_paragraphs(nested_table)


def _is_field_paragraph(paragraph) -> bool:
    """Return True when the paragraph contains a Word field like TOC."""
    try:
        xml = paragraph._p.xml
    except Exception:
        return False
    return "w:instrText" in xml or "w:fldChar" in xml


def _contains_image(paragraph) -> bool:
    """Return True if the paragraph has an inline/floating image (drawing) or a VML picture."""
    p_xml = paragraph._p
    drawing_tags = (
        qn("w:drawing"),
        "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline",
        "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor",
        "{urn:schemas-microsoft-com:vml}shape",
        "{urn:schemas-microsoft-com:vml}imagedata",
    )
    for tag in drawing_tags:
        if p_xml.find(f".//{tag}") is not None:
            return True
    return False


def _clean_empty_paragraphs(document: Document) -> None:
    """Rapikan paragraf kosong beruntun jadi maksimal SATU baris kosong saja.

    Paragraf berisi gambar/logo atau field Word (TOC, page number) TIDAK
    pernah dihapus. Baris kosong TUNGGAL antar bagian dibiarkan (spasi
    visual yang wajar); yang dirapikan hanya runtutan 2+ baris kosong
    berturut-turut ("spasinya banyak") supaya tidak ada jarak berlebihan.
    """
    paragraphs = list(_iter_all_paragraphs(document))
    previous_was_blank = False
    for paragraph in paragraphs:
        if _is_field_paragraph(paragraph):
            previous_was_blank = False
            continue
        if _contains_image(paragraph):
            previous_was_blank = False
            continue

        is_blank = not paragraph.text or not paragraph.text.strip()
        if is_blank and previous_was_blank:
            p_element = paragraph._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
            continue

        previous_was_blank = is_blank


def _insert_custom_header(document: Document, config: dict) -> None:
    """Insert a custom header into the first section only."""
    header_text = config.get("header_text")
    if not header_text:
        return

    first_section = document.sections[0]
    header = first_section.header
    header.is_linked_to_previous = False
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_para.text = header_text
    header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    header_font_name = config.get("header_font_name")
    header_font_size_pt = config.get("header_font_size_pt")
    if header_para.runs:
        run = header_para.runs[0]
    else:
        run = header_para.add_run(header_text)

    if header_font_name:
        run.font.name = header_font_name
    if header_font_size_pt is not None:
        run.font.size = Pt(header_font_size_pt)


def _inject_dynamic_page_numbering(document: Document) -> None:
    """Add automatic page numbering footers and reset numbering per section.

    Dokumen final bisa punya BANYAK section (bukan cuma 1-2): tiap bab baru
    disisipkan sebagai section break tersendiri (lihat mailmerge.py) supaya
    tiap bab mulai di halaman ganjil baru. Sebelumnya fungsi ini cuma
    menangani section pertama & kedua secara eksplisit — section ketiga dan
    seterusnya (bab ke-2 dst.) tidak pernah diberi nomor halaman ("belum
    berjalan"). Sekarang SEMUA section diproses:
      - Section pertama = halaman depan (cover, identitas, kata pengantar,
        daftar isi) -> angka romawi kecil (i, ii, iii, ...).
      - Section kedua dan seterusnya (isi naskah per-bab + tentang penulis)
        -> angka arab, dimulai dari 1 HANYA di section kedua; section-section
        setelahnya TIDAK di-restart supaya penomoran berlanjut wajar dan
        tidak "loncat balik ke 1" / dobel di tiap pergantian bab.
    """
    sections = document.sections

    if len(sections) == 0:
        return

    if len(sections) == 1:
        _set_section_page_number_type(sections[0], fmt="decimal", start=1)
        _set_footer_page_number(sections[0].footer)
        return

    # Section pertama: halaman depan, angka romawi kecil.
    _set_section_page_number_type(sections[0], fmt="lowerRoman", start=1)
    _set_footer_page_number(sections[0].footer)

    # Section kedua dst.: angka arab. Restart ke 1 cuma di section kedua
    # (awal isi naskah); section berikutnya melanjutkan penomoran otomatis
    # (start=None -> atribut w:start tidak dipasang sama sekali).
    for idx in range(1, len(sections)):
        start = 1 if idx == 1 else None
        _set_section_page_number_type(sections[idx], fmt="decimal", start=start)
        _set_footer_page_number(sections[idx].footer)


def _count_section_nodes(document: Document) -> int:
    """Count w:sectPr nodes in the document body XML for diagnostic purposes."""
    return len(document.element.body.findall(qn("w:sectPr")))



# Urutan resmi elemen anak <w:sectPr> menurut skema CT_SectPr (ECMA-376).
# w:pgNumType WAJIB muncul SEBELUM elemen-elemen berikut ini di dalam sectPr.
# Menaruhnya setelah salah satu dari elemen ini (mis. dengan sect_pr.append()
# begitu saja) melanggar urutan skema -> Word menolak membuka file / minta
# "repair", walaupun parser yang lebih longgar seperti LibreOffice tetap mau
# membukanya tanpa keluhan.
_SECTPR_ELEMENTS_AFTER_PGNUMTYPE = (
    "w:cols",
    "w:formProt",
    "w:vAlign",
    "w:noEndnote",
    "w:titlePg",
    "w:textDirection",
    "w:bidi",
    "w:rtlGutter",
    "w:docGrid",
    "w:printerSettings",
    "w:sectPrChange",
)


def _set_section_page_number_type(section, fmt: str, start: int | None) -> None:
    """Set the page number format for a section.

    Kalau `start` diisi (int), section ini akan RESTART penomoran mulai dari
    angka tersebut. Kalau `start` None, atribut w:start sengaja TIDAK
    dipasang sama sekali sehingga Word melanjutkan penomoran dari section
    sebelumnya (dipakai untuk bab ke-2 dst. supaya tidak restart/dobel).
    """
    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:pgNumType"))
    if existing is not None:
        sect_pr.remove(existing)

    pg_num_type = OxmlElement("w:pgNumType")
    pg_num_type.set(qn("w:fmt"), fmt)
    if start is not None:
        pg_num_type.set(qn("w:start"), str(start))

    # Cari elemen anak pertama yang seharusnya berada SETELAH pgNumType
    # (mis. w:cols, w:titlePg, w:docGrid), lalu sisipkan pgNumType tepat
    # sebelum elemen tersebut supaya urutan skema tetap valid.
    insert_before = None
    for child in sect_pr:
        if qn_localname(child.tag) in _SECTPR_ELEMENTS_AFTER_PGNUMTYPE:
            insert_before = child
            break

    if insert_before is not None:
        insert_before.addprevious(pg_num_type)
    else:
        sect_pr.append(pg_num_type)

    try:
        section.footer.is_linked_to_previous = False
    except Exception:
        pass


def qn_localname(tag: str) -> str:
    """Ubah Clark-notation tag ('{namespace}local') balik jadi 'w:local' untuk dibandingkan."""
    if "}" not in tag:
        return tag
    local = tag.split("}", 1)[1]
    return f"w:{local}"


def _set_footer_page_number(footer) -> None:
    """Insert a centered PAGE field into the footer using OXML."""
    footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_paragraph.clear()
    footer_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    run = footer_paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr_text)
    run._r.append(fld_separate)
    run._r.append(fld_end)


if __name__ == "__main__":
    sample_config = {
        "margin_top_cm": 2.0,
        "margin_bottom_cm": 2.0,
        "margin_left_cm": 3.0,
        "margin_right_cm": 3.0,
        "font_name": "Times New Roman",
        "font_size_pt": 12,
        "line_spacing": 1.15,
        "alignment": "justify",
        "header_text": "NAMA DOKUMEN - DIFORMAT OTOMATIS",
        "header_font_name": "Arial",
        "header_font_size_pt": 9,
    }

    normalize_docx_layout(
        input_docx_path="input_messy.docx",
        output_docx_path="output_normalized.docx",
        config=sample_config,
    )