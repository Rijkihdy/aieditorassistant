"""
Mail merge dinamis & auto-formatting untuk template buku Guepedia (.docx).
"""
import copy
import io
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

from docx_formatter import normalize_docx_layout
from naskah_parser import split_into_sections, strip_front_matter

CHAPTER_TITLE_PLACEHOLDER = "Judul Bab"
CHAPTER_BODY_PLACEHOLDER = "(Isi naskah)"


class MailMergeError(Exception):
    """Dilempar jika template tidak punya placeholder yang diharapkan."""


def generate_book_docx(
    template_path: str,
    output_path: str,
    nama_penulis: str,
    judul_naskah: str,
    isbn: str,
    tahun_cetak: str,
    kata_pengantar_text: str = "",
    naskah_text: str = "",
    chapter_title: str = "",
    format_config: dict | None = None,
    source_docx_bytes: bytes | None = None,
    qrcbn: str = "",
    sinopsis_text: str = "",
) -> None:
    """Isi template buku dengan data dinamis, lalu terapkan format dokumen dinamis ke output akhir."""
    document = Document(template_path)

    # 1. Replacement Halaman Depan & Redaksi
    simple_replacements = {
        "Nama Penulis": nama_penulis or "Nama Penulis",
        "Judul Naskah": judul_naskah or "Judul Naskah",
        "ISBN: ": f"ISBN: {isbn}" if isbn else "ISBN: ",
        "Februari 2023": tahun_cetak or "Mei 2025",
    }

    for paragraph in _iter_all_paragraphs(document):
        _apply_simple_replacements(paragraph, simple_replacements)

    if qrcbn:
        _insert_qrcbn(document, qrcbn)

    # 2. Siapkan baris naskah, buang front matter buatan penulis sendiri
    #    (Kata Pengantar & listing Daftar Isi) karena dokumen final sudah
    #    punya slotnya sendiri (placeholder Kata Pengantar + field TOC otomatis).
    if source_docx_bytes is not None:
        source_doc = Document(io.BytesIO(source_docx_bytes))
        raw_lines = [p.text.strip() for p in source_doc.paragraphs if p.text and p.text.strip()]
    else:
        raw_lines = [line.strip() for line in naskah_text.splitlines() if line.strip()]

    auto_kata_pengantar, body_lines = strip_front_matter(raw_lines) if raw_lines else ("", [])
    sections = split_into_sections(body_lines, fallback_title=chapter_title) if body_lines else []

    # 3. Isi Sinopsis (opsional — dari input manual/upload ATAU hasil generate AI di app.py)
    if sinopsis_text and sinopsis_text.strip():
        _insert_sinopsis(document, sinopsis_text)

    # 4. Isi Kata Pengantar — pakai input manual dari form kalau diisi,
    #    kalau kosong pakai yang otomatis terdeteksi dari naskah asli.
    final_kata_pengantar = kata_pengantar_text or auto_kata_pengantar
    if final_kata_pengantar:
        _replace_kata_pengantar(document, final_kata_pengantar)

    # 5. Masukkan Isi Naskah Multi-Bab (Daftar Isi bawaan naskah sudah dibuang),
    #    tiap bab baru dimulai di halaman ganjil baru (section break oddPage).
    _insert_manuscript_sections(document, sections, format_config=format_config)

    _append_toc_field(document)
    _ensure_update_fields(document)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)

    # 6. Terapkan Format Dinamis Global (Margin, Font, Header, Spacing, dll)
    if format_config:
        normalize_docx_layout(output_path, output_path, format_config)


def _iter_all_paragraphs(document: Document):
    for paragraph in document.paragraphs:
        yield paragraph

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph

    for section in document.sections:
        containers = [
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ]
        for container in containers:
            if container is None:
                continue
            for paragraph in container.paragraphs:
                yield paragraph


def _apply_simple_replacements(paragraph: Paragraph, replacements: dict) -> None:
    text = paragraph.text
    if not text:
        return

    new_text = text
    for old, new in replacements.items():
        if old in new_text:
            new_text = new_text.replace(old, new)

    if new_text != text:
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""


def _insert_qrcbn(document: Document, qrcbn: str) -> None:
    """Sisipkan baris QRCBN tepat setelah baris ISBN di halaman francis."""
    for i, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip().startswith("ISBN:"):
            anchor_element = paragraph._p
            new_p = document.add_paragraph()
            anchor_element.addnext(new_p._p)
            new_p.text = f"QRCBN: {qrcbn}"
            for run in paragraph.runs:
                # samakan format (font/size) dengan baris ISBN di atasnya
                if run.text.strip():
                    for new_run in new_p.runs:
                        new_run.font.name = run.font.name
                        new_run.font.size = run.font.size
                    break
            return


def _insert_sinopsis(document: Document, sinopsis_text: str) -> None:
    """Sisipkan halaman SINOPSIS baru sebelum halaman Kata Pengantar."""
    anchor_paragraph = None
    for paragraph in document.paragraphs:
        if "KATA PENGANTAR" in paragraph.text:
            anchor_paragraph = paragraph
            break

    if anchor_paragraph is None:
        anchor_paragraph = document.paragraphs[-1]

    lines = [_normalize_spacing(line) for line in sinopsis_text.splitlines() if line.strip()]
    if not lines:
        return

    anchor_element = anchor_paragraph._p

    heading_p = document.add_paragraph()
    anchor_element.addprevious(heading_p._p)
    heading_p.text = "SINOPSIS"
    try:
        heading_p.style = document.styles["Heading 1"]
    except KeyError:
        pass
    heading_p.paragraph_format.page_break_before = True

    insert_after = heading_p._p
    for line in lines:
        body_p = document.add_paragraph()
        insert_after.addnext(body_p._p)
        insert_after = body_p._p
        body_p.text = line
        body_p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        body_p.paragraph_format.first_line_indent = Cm(0.75)

    # Kata Pengantar tetap mulai di halaman baru setelah Sinopsis
    anchor_paragraph.paragraph_format.page_break_before = True


def _normalize_spacing(text: str) -> str:
    """Rapikan spasi ganda/berlebih jadi satu spasi, buang spasi awal-akhir."""
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def _replace_kata_pengantar(document: Document, kata_pengantar_text: str) -> None:
    """Isi bagian Kata Pengantar sebelum Bab 1."""
    lines = [_normalize_spacing(line) for line in kata_pengantar_text.splitlines() if line.strip()]
    if not lines:
        return

    for i, paragraph in enumerate(document.paragraphs):
        if "KATA PENGANTAR" in paragraph.text:
            if i + 1 < len(document.paragraphs):
                target = document.paragraphs[i + 1]
                target.text = lines[0]
                anchor_element = target._p
                parent = target._parent
                for line in lines[1:]:
                    new_element = copy.deepcopy(anchor_element)
                    anchor_element.addnext(new_element)
                    anchor_element = new_element
                    new_paragraph = Paragraph(new_element, parent)
                    new_paragraph.text = line
            break


def _insert_manuscript_sections(
    document: Document,
    sections: list[tuple[str, list[str]]],
    format_config: dict | None = None,
) -> None:
    """Sisipkan sections (judul_bab, isi_baris) hasil naskah_parser ke placeholder bab, dengan styling dinamis."""
    target_p = None
    for paragraph in document.paragraphs:
        if CHAPTER_BODY_PLACEHOLDER in paragraph.text or CHAPTER_TITLE_PLACEHOLDER in paragraph.text:
            target_p = paragraph
            break

    if target_p is None:
        target_p = document.add_paragraph()

    if not sections:
        p_element = target_p._element
        if p_element.getparent() is not None:
            p_element.getparent().remove(p_element)
        return

    # Ambil setting font & line spacing dari format_config dinamis
    font_name = format_config.get("font_name", "Times New Roman") if format_config else "Times New Roman"
    font_size = format_config.get("font_size_pt", 12) if format_config else 12
    line_spacing = format_config.get("line_spacing", 1.15) if format_config else 1.15
    heading_alignment_key = (format_config.get("heading_alignment", "left") if format_config else "left")
    heading_alignment = _ALIGNMENT_MAP.get(heading_alignment_key.lower(), WD_PARAGRAPH_ALIGNMENT.LEFT)

    # sectPr dasar dipakai sebagai cetakan section break tiap awal bab
    # (disalin dari section terakhir template supaya margin/header/footer konsisten)
    base_sect_pr = document.element.body.find(qn("w:sectPr"))

    anchor_element = target_p._p

    def _new_paragraph_after(anchor_el):
        """Buat paragraf baru lewat API python-docx (murah), lalu pindahkan
        tepat setelah anchor_el. Jauh lebih hemat memori/CPU dibanding
        copy.deepcopy per baris, yang mahal untuk naskah ratusan halaman."""
        p = document.add_paragraph()
        anchor_el.addnext(p._p)
        return p

    for section_title, body_lines in sections:
        if section_title:
            is_bab = bool(re.match(r"^Bab\s+\d+", section_title, re.IGNORECASE))

            if is_bab and base_sect_pr is not None:
                # Section break "mulai halaman ganjil baru" tepat sebelum judul bab.
                marker_p = _new_paragraph_after(anchor_element)
                anchor_element = marker_p._p
                marker_p_pr = marker_p._p.get_or_add_pPr()
                marker_p_pr.append(_make_odd_page_sectpr(base_sect_pr))

            p_new = _new_paragraph_after(anchor_element)
            anchor_element = p_new._p

            p_new.text = _normalize_spacing(section_title)

            if is_bab:
                try:
                    p_new.style = document.styles["Heading 1"]
                except KeyError:
                    p_new.style = document.styles["Normal"]
                p_new.alignment = heading_alignment
                p_new.paragraph_format.space_before = Pt(18)
                p_new.paragraph_format.space_after = Pt(12)
                heading_font_size = font_size + 4
            elif section_title.startswith("[") and section_title.endswith("]"):
                try:
                    p_new.style = document.styles["Heading 2"]
                except KeyError:
                    p_new.style = document.styles["Normal"]
                p_new.alignment = heading_alignment
                p_new.paragraph_format.space_before = Pt(12)
                p_new.paragraph_format.space_after = Pt(4)
                heading_font_size = font_size + 2
            else:
                p_new.style = document.styles["Normal"]
                p_new.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                heading_font_size = font_size

            for run in p_new.runs:
                run.font.name = font_name
                run.font.size = Pt(heading_font_size)
                if section_title.startswith("[") or is_bab:
                    run.font.bold = True

        for line in body_lines:
            clean_line = _normalize_spacing(line)
            if not clean_line:
                continue

            p_new = _new_paragraph_after(anchor_element)
            anchor_element = p_new._p

            p_new.text = clean_line

            # Paragraf Isi Naskah biasa (Format Dinamis)
            p_new.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            p_new.paragraph_format.first_line_indent = Cm(0.75)
            p_new.paragraph_format.line_spacing = line_spacing

            for run in p_new.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)

    # Hapus paragraf placeholder bawaan
    p_element = target_p._element
    if p_element.getparent() is not None:
        p_element.getparent().remove(p_element)


_ALIGNMENT_MAP = {
    "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
    "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
    "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
}

# Elemen sectPr yang menurut skema CT_SectPr WAJIB muncul SEBELUM w:type.
_SECTPR_LOCALNAMES_BEFORE_TYPE = {"headerReference", "footerReference", "footnotePr", "endnotePr"}


def _make_odd_page_sectpr(base_sect_pr) -> "OxmlElement":
    """Salin sectPr dasar & set w:type jadi 'oddPage' — dipakai sebagai
    penanda section break di w:pPr suatu paragraf, supaya bab berikutnya
    otomatis mulai di halaman ganjil baru (Word akan menyisipkan halaman
    kosong sendiri kalau perlu)."""
    new_sect_pr = copy.deepcopy(base_sect_pr)

    # Buang titlePg: kalau ikut disalin, halaman pertama SETIAP bab akan
    # pakai header/footer "halaman judul" (biasanya tanpa nomor halaman),
    # padahal kita mau nomor halaman tetap konsisten di halaman pembuka bab.
    for title_pg in new_sect_pr.findall(qn("w:titlePg")):
        new_sect_pr.remove(title_pg)

    for existing_type in new_sect_pr.findall(qn("w:type")):
        new_sect_pr.remove(existing_type)

    type_el = OxmlElement("w:type")
    type_el.set(qn("w:val"), "oddPage")

    insert_after = None
    for child in new_sect_pr:
        if _local_tag(child.tag) in _SECTPR_LOCALNAMES_BEFORE_TYPE:
            insert_after = child
    if insert_after is not None:
        insert_after.addnext(type_el)
    else:
        new_sect_pr.insert(0, type_el)

    return new_sect_pr


def _append_toc_field(document: Document) -> None:
    """Tambah field TOC ke dokumen agar Word bisa menampilkan daftar isi otomatis.

    PENTING: w:fldChar / w:instrText harus jadi anak <w:r> (run), BUKAN anak
    langsung <w:p> (paragraf) — kalau tidak, Word menganggap file corrupt dan
    minta "repair" saat dibuka. Ikuti pola begin -> separate -> end yang sama
    dengan _set_footer_page_number di docx_formatter.py.
    """
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run()

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_char_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instr_text)

    fld_char_separate = OxmlElement("w:fldChar")
    fld_char_separate.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_char_separate)

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_end)


def _local_tag(tag: str) -> str:
    """Ubah Clark-notation tag ('{namespace}local') jadi nama lokalnya saja."""
    return tag.split("}", 1)[1] if "}" in tag else tag


# Nama elemen (localname, tanpa prefix) yang menurut skema CT_Settings WAJIB
# muncul SETELAH w:updateFields. Dipakai untuk mencari titik sisip yang benar
# supaya urutan skema tidak dilanggar (append() mentah ke akhir <w:settings>
# bisa menaruh updateFields SETELAH elemen seperti w:compat/w:rsids/mathPr,
# yang membuat Word menganggap file corrupt).
_SETTINGS_LOCALNAMES_AFTER_UPDATEFIELDS = {
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
    "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
    "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
    "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
    "doNotDemarcateInvalidXml", "doNotValidateAgainstSchema",
}


def _ensure_update_fields(document: Document) -> None:
    """Pastikan Word diperintahkan meng-update field TOC saat dokumen dibuka."""
    settings_element = document.settings.element
    existing = settings_element.find(qn("w:updateFields"))
    if existing is not None:
        return

    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "1")

    insert_before = None
    for child in settings_element:
        if _local_tag(child.tag) in _SETTINGS_LOCALNAMES_AFTER_UPDATEFIELDS:
            insert_before = child
            break

    if insert_before is not None:
        insert_before.addprevious(update_fields)
    else:
        settings_element.append(update_fields)