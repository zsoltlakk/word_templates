from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        tblBorders.append(border)
    tblPr.append(tblBorders)


def set_cell_font(cell, size_pt=8, bold=False):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size_pt)
            run.font.name = "Arial"
            run.bold = bold


def add_run_to_cell(cell, text, size_pt=8, bold=False, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment
    run = paragraph.add_run(text)
    run.font.size = Pt(size_pt)
    run.font.name = "Arial"
    run.bold = bold
    return run


doc = Document()

# Page setup: landscape A4
section = doc.sections[0]
section.page_width = Cm(29.7)
section.page_height = Cm(21.0)
section.left_margin = Cm(1.0)
section.right_margin = Cm(1.0)
section.top_margin = Cm(1.0)
section.bottom_margin = Cm(1.0)

# Default font
style = doc.styles["Normal"]
font = style.font
font.name = "Arial"
font.size = Pt(8)

# ── Header table (2 rows × 3 cols, no border) ───────────────────────────────
header_table = doc.add_table(rows=2, cols=3)
header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(header_table)

# Row 0
add_run_to_cell(header_table.cell(0, 0), "{{LOGO}}", size_pt=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
add_run_to_cell(header_table.cell(0, 1), "Folyószámla összesítő", size_pt=14, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_run_to_cell(header_table.cell(0, 2), "Dátum: {{DATUM}}", size_pt=8, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

# Row 1
header_table.cell(1, 0).paragraphs[0].text = ""
add_run_to_cell(header_table.cell(1, 1), "{{EV}}. {{HONAP}}", size_pt=12, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_run_to_cell(header_table.cell(1, 2), "Oldal: {{OLDAL}}", size_pt=8, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

doc.add_paragraph("")

# ── Main data table (Table Grid style) ──────────────────────────────────────
cols_count = 11
main_table = doc.add_table(rows=3 + 20 + 1, cols=cols_count)
main_table.style = "Table Grid"
main_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Column widths
col_widths = [
    Cm(2.5),  # Banki dátum
    Cm(2.2),  # Számla bevét
    Cm(2.2),  # Áttutóra bevét
    Cm(2.2),  # Egyéb
    Cm(2.2),  # Számláról túlfiz visszautalás
    Cm(2.2),  # Áttutóról visszautalás
    Cm(2.2),  # HM VGH jogi költség
    Cm(2.2),  # HM EI jogi költség
    Cm(2.2),  # Banki kezelési költség
    Cm(2.2),  # Egyenleg leemelés
    Cm(3.5),  # Záróegyenleg
]

def set_row_widths(row):
    for idx, width in enumerate(col_widths):
        row.cells[idx].width = width

# ── Row 0: Group headers ─────────────────────────────────────────────────────
row0 = main_table.rows[0]
set_row_widths(row0)

# Col 0: empty
row0.cells[0].paragraphs[0].text = ""

# Cols 1-3: Jóváírások
row0.cells[1].merge(row0.cells[3])
add_run_to_cell(row0.cells[1], "Jóváírások", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Cols 4-5: Terhelések
row0.cells[4].merge(row0.cells[5])
add_run_to_cell(row0.cells[4], "Terhelések", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Cols 6-7: Jogi költségek
row0.cells[6].merge(row0.cells[7])
add_run_to_cell(row0.cells[6], "Jogi költségek", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Col 8: Banki kezelési költség (standalone header)
add_run_to_cell(row0.cells[8], "Banki kezelési költség", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Col 9: Egyenleg leemelés (standalone header)
add_run_to_cell(row0.cells[9], "Egyenleg leemelés", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Col 10: Záróegyenleg (standalone header)
add_run_to_cell(row0.cells[10], "Záróegyenleg", size_pt=8, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# ── Row 1: Column sub-headers ─────────────────────────────────────────────────
column_headers = [
    "Banki dátum",
    "Számla bevét",
    "Áttutóra bevét",
    "Egyéb",
    "Számláról túlfiz\nvisszautalás",
    "Áttutóról\nvisszautalás",
    "HM VGH jogi\nköltség",
    "HM EI jogi\nköltség",
    "Banki kezelési\nköltség",
    "Egyenleg\nleemelés",
    "Záróegyenleg",
]

row1 = main_table.rows[1]
set_row_widths(row1)
for idx, header_text in enumerate(column_headers):
    add_run_to_cell(row1.cells[idx], header_text, size_pt=7, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

# ── Row 2: Previous month closing balance ────────────────────────────────────
row2 = main_table.rows[2]
set_row_widths(row2)
# Cols 0-9 empty
for idx in range(10):
    row2.cells[idx].paragraphs[0].text = ""
# Col 10: previous balance placeholder
p = row2.cells[10].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run("{{EV_ELOZO}}. {{HONAP_ELOZO}}. havi záróegyenleg:\n{{ELOZO_HAVI_ZAROEGYENLEG}}")
run.font.size = Pt(7)
run.font.name = "Arial"

# ── Rows 3-22: Placeholder data rows ─────────────────────────────────────────
for n in range(1, 21):
    row = main_table.rows[2 + n]
    set_row_widths(row)
    placeholders = [
        (f"{{{{BANKI_DATUM_{n}}}}}", WD_ALIGN_PARAGRAPH.LEFT),
        (f"{{{{SZAMLA_BEVET_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{ATTUTORA_BEVET_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{EGYEB_JOVAIRAS_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{SZAMLAROL_TULFIZ_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{ATTUTOR_VISSZAUT_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{HM_VGH_JOGI_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{HM_EI_JOGI_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{BANKI_KEZ_KOLTSEG_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{EGYENLEG_LEEMELES_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
        (f"{{{{ZAROEGYENLEG_{n}}}}}", WD_ALIGN_PARAGRAPH.RIGHT),
    ]
    for idx, (placeholder_text, align) in enumerate(placeholders):
        p = row.cells[idx].paragraphs[0]
        p.alignment = align
        run = p.add_run(placeholder_text)
        run.font.size = Pt(7)
        run.font.name = "Arial"

# ── Last row: merged comment/note placeholder ─────────────────────────────────
note_row = main_table.rows[23]
set_row_widths(note_row)
note_row.cells[0].merge(note_row.cells[10])
p = note_row.cells[0].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p.add_run("{{MEGJEGYZES}}")
run.font.size = Pt(7)
run.font.name = "Arial"

# ── Save ──────────────────────────────────────────────────────────────────────
filename = "folyoszamla_osszesito_template.docx"
doc.save(filename)
print(f"Template saved as: {filename}")
