import os
import tempfile
import subprocess
import shutil
import logging
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger(__name__)

# ─── Colors ───
BLUE  = RGBColor(0x1A, 0x52, 0x76)
GRAY  = RGBColor(0x5D, 0x6D, 0x7E)
BLACK = RGBColor(0x1C, 0x1C, 0x1C)
BLUE_HEX = "1A5276"


# ════════════════════════════════════════
#              HELPERS
# ════════════════════════════════════════

def font(run, size, color=None, bold=False, italic=False):
    """Apply font styling to a run"""
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic


def add_right_tab(para, position=9026):
    """Add a right-aligned tab stop to paragraph"""
    pPr = para._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(position))
    tabs.append(tab)
    pPr.append(tabs)


def add_section_divider(doc, title):
    """Section title with blue bottom border"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(14)
    para.paragraph_format.space_after = Pt(4)

    run = para.add_run(title.upper())
    font(run, 12, BLUE, bold=True)

    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), BLUE_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)


def remove_cell_borders(cell):
    """Remove all borders from a table cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:color'), 'auto')
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_job_entry(doc, title, company, date_range, bullets):
    """Add a work experience block"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    add_right_tab(p)
    font(p.add_run(title), 11, BLACK, bold=True)
    p.add_run('\t').font.name = 'Arial'
    font(p.add_run(date_range), 10, GRAY)

    pc = doc.add_paragraph()
    pc.paragraph_format.space_before = Pt(0)
    pc.paragraph_format.space_after = Pt(4)
    font(pc.add_run(company), 10, BLUE, italic=True)

    for b in bullets:
        pb = doc.add_paragraph(style='List Bullet')
        pb.paragraph_format.space_before = Pt(2)
        pb.paragraph_format.space_after = Pt(2)
        font(pb.add_run(b), 10, BLACK)


def add_skills_table(doc, skills):
    """Two-column borderless skills table"""
    half = (len(skills) + 1) // 2
    left  = skills[:half]
    right = skills[half:]

    table = doc.add_table(rows=half, cols=2)

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
    tblPr.append(tblBorders)

    for i in range(half):
        row = table.rows[i]

        cl = row.cells[0]
        cl.width = Twips(4513)
        remove_cell_borders(cl)
        pl = cl.paragraphs[0]
        pl.clear()
        font(pl.add_run("▸ "), 10, BLUE)
        font(pl.add_run(left[i] if i < len(left) else ""), 10, BLACK)

        cr = row.cells[1]
        cr.width = Twips(4513)
        remove_cell_borders(cr)
        if i < len(right):
            pr = cr.paragraphs[0]
            pr.clear()
            font(pr.add_run("▸ "), 10, BLUE)
            font(pr.add_run(right[i]), 10, BLACK)


# ════════════════════════════════════════
#           GENERATE CV (DOCX)
# ════════════════════════════════════════

def generate_cv(data: dict) -> str:
    """
    Build a DOCX CV from user data.
    Returns path to the generated temp file.
    Raises exception on failure.
    """
    tmp_path = None
    try:
        doc = Document()

        # Page margins (A4)
        for section in doc.sections:
            section.top_margin    = Cm(1.5)
            section.bottom_margin = Cm(1.5)
            section.left_margin   = Cm(1.8)
            section.right_margin  = Cm(1.8)

        # ── HEADER ──────────────────────────
        p_name = doc.add_paragraph()
        p_name.paragraph_format.space_after = Pt(2)
        font(p_name.add_run(data['name']), 26, BLUE, bold=True)

        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_after = Pt(6)
        font(p_title.add_run(data['job_title']), 12, GRAY)

        contact_parts = [
            f"📧 {data['email']}",
            f"📞 {data['phone']}",
            f"📍 {data['location']}",
        ]
        if data.get('linkedin'):
            contact_parts.append(f"🔗 {data['linkedin']}")
        if data.get('github'):
            contact_parts.append(f"💻 {data['github']}")

        p_contact = doc.add_paragraph()
        p_contact.paragraph_format.space_after = Pt(12)
        font(p_contact.add_run("   |   ".join(contact_parts)), 9, GRAY)

        # ── SUMMARY ─────────────────────────
        add_section_divider(doc, "Professional Summary")
        p_sum = doc.add_paragraph()
        p_sum.paragraph_format.space_before = Pt(4)
        p_sum.paragraph_format.space_after  = Pt(10)
        font(p_sum.add_run(data['summary']), 10, BLACK)

        # ── EXPERIENCE ──────────────────────
        add_section_divider(doc, "Experience")
        for exp in data['experiences']:
            add_job_entry(doc, exp['title'], exp['company'], exp['date'], exp['bullets'])

        # ── EDUCATION ───────────────────────
        add_section_divider(doc, "Education")
        p_edu = doc.add_paragraph()
        p_edu.paragraph_format.space_before = Pt(8)
        p_edu.paragraph_format.space_after  = Pt(2)
        add_right_tab(p_edu)
        font(p_edu.add_run(data['edu_degree']), 11, BLACK, bold=True)
        p_edu.add_run('\t').font.name = 'Arial'
        font(p_edu.add_run(data['edu_date']), 10, GRAY)

        p_uni = doc.add_paragraph()
        p_uni.paragraph_format.space_before = Pt(0)
        p_uni.paragraph_format.space_after  = Pt(10)
        font(p_uni.add_run(data['edu_university']), 10, GRAY, italic=True)

        # ── SKILLS ──────────────────────────
        add_section_divider(doc, "Skills")
        doc.add_paragraph().paragraph_format.space_before = Pt(6)
        add_skills_table(doc, data['skills'])

        # ── CERTIFICATES ────────────────────
        if data.get('certificates'):
            add_section_divider(doc, "Certificates")
            for cert in data['certificates']:
                p_cert = doc.add_paragraph()
                p_cert.paragraph_format.space_before = Pt(4)
                p_cert.paragraph_format.space_after  = Pt(2)
                add_right_tab(p_cert)
                font(p_cert.add_run(cert['name']), 10, BLACK, bold=True)
                font(p_cert.add_run(' — '), 10, GRAY)
                font(p_cert.add_run(cert['issuer']), 10, GRAY, italic=True)
                p_cert.add_run('\t').font.name = 'Arial'
                font(p_cert.add_run(cert['date']), 10, GRAY)

        # ── LANGUAGES ───────────────────────
        add_section_divider(doc, "Languages")
        p_lang = doc.add_paragraph()
        p_lang.paragraph_format.space_before = Pt(6)
        font(p_lang.add_run(data['languages']), 10, BLACK)

        # Save to temp file
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        tmp_path = tmp.name
        tmp.close()
        doc.save(tmp_path)
        logger.info(f"CV generated: {tmp_path}")
        return tmp_path

    except Exception as e:
        # Clean up on failure
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
        logger.error(f"CV generation failed: {e}")
        raise


# ════════════════════════════════════════
#           CONVERT TO PDF
# ════════════════════════════════════════

def _find_libreoffice() -> str:
    """Find LibreOffice binary across different OS paths."""
    candidates = [
        'libreoffice',
        'soffice',
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]
    for candidate in candidates:
        if shutil.which(candidate) or os.path.exists(candidate):
            return candidate
    raise FileNotFoundError(
        "LibreOffice not found. Install it or use Word format instead."
    )


def convert_to_pdf(docx_path: str) -> str:
    """
    Convert DOCX to PDF using LibreOffice.
    Returns path to the generated PDF file.
    Raises exception with clear message on failure.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        lo_bin = _find_libreoffice()

        result = subprocess.run(
            [lo_bin, '--headless', '--convert-to', 'pdf',
             '--outdir', tmp_dir, docx_path],
            capture_output=True,
            text=True,
            timeout=60
        )

        if result.returncode != 0:
            raise Exception(f"LibreOffice error: {result.stderr.strip()}")

        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path  = os.path.join(tmp_dir, base_name + '.pdf')

        if not os.path.exists(pdf_path):
            raise Exception("PDF file not found after conversion")

        logger.info(f"PDF converted: {pdf_path}")
        return pdf_path

    except Exception:
        # Clean up temp dir on failure
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise
