#!/usr/bin/env python3
"""
ReportTerm5 - Automated Cooperative Training Report Generator
Generates the complete Final Cooperative Training Report for the "Riaya" Project
at the University of Hail (College of Computer Science and Engineering),
tailored for Shadia Essa Al-Enezi at Al-Hait General Hospital.

Adheres strictly to the gold template guidelines in ReportLMM / ReportTerm5:
- Times New Roman 12pt, 1.5 line spacing, 10pt space after
- Headings: H1 (16pt bold #1B3A5C), H2 (14pt bold #1B3A5C), H3 (13pt bold #2E5D8A)
- Strict XML SEQ fields inside <w:r> tags for Figure and Table captions
- Table captions above tables, Figure captions below images
- Centered 6.27-inch images matching printable page margins
- TOC, Figure TOC, and Table TOC automated fields
- 100% compliant XML with zero broken fields
"""

import os
import sys
import zipfile
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from lxml import etree

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Paths
BASE_DIR = r'C:\MAMP\htdocs\Riaya'
REPO_DIR = os.path.join(BASE_DIR, 'ReportTerm5')
TEMPLATE_DOCX = os.path.join(REPO_DIR, 'template.docx')
OUTPUT_DIR = os.path.join(REPO_DIR, 'FinalReportsDOCX')
OUTPUT_DOCX = os.path.join(OUTPUT_DIR, 'Riaya_Final_Report.docx')
REPORT_EXPORT_DOCX = os.path.join(BASE_DIR, 'Report', 'Riaya_Final_Report.docx')

# Media directories
UOH_LOGO = os.path.join(BASE_DIR, 'Report', 'template_media', 'word', 'media', 'image1.png')
ORG_MEDIA = os.path.join(BASE_DIR, 'Report', 'training_org_media')
DIAG_DIR = os.path.join(BASE_DIR, 'docs', 'diagrams')
SCREEN_DIR = os.path.join(BASE_DIR, 'Report', 'converted_screenshots')

# Colors
COLOR_DARK_BLUE = RGBColor(0x1B, 0x3A, 0x5C)
COLOR_MED_BLUE = RGBColor(0x2E, 0x5D, 0x8A)
HEX_NAVY = "1B3A5C"
HEX_LIGHT_ROW = "F9FBFC"
HEX_BORDER = "1B3A5C"

def set_run_font(run, name="Times New Roman", size_pt=12, bold=False, italic=False, color=None):
    run.font.name = name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    # Ensure East Asia and Complex Script match
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:cs'), name)

def set_para_format(p, before_pt=0, after_pt=10, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    p.paragraph_format.space_before = Pt(before_pt)
    p.paragraph_format.space_after = Pt(after_pt)
    p.paragraph_format.line_spacing = line_spacing
    p.alignment = align

def add_paragraph_styled(doc, text="", bold_prefix=None, before_pt=0, after_pt=10, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    set_para_format(p, before_pt, after_pt, line_spacing, align)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        set_run_font(r_pre, "Times New Roman", 12, bold=True, italic=False, color=RGBColor(0, 0, 0))
    if text:
        r_txt = p.add_run(text)
        set_run_font(r_txt, "Times New Roman", 12, bold=False, italic=False, color=RGBColor(0, 0, 0))
    return p

def add_h1(doc, text):
    p = doc.add_paragraph(style='Heading 1')
    set_para_format(p, before_pt=16, after_pt=8, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.text = ""
    r = p.add_run(text)
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)
    return p

def add_h2(doc, text):
    p = doc.add_paragraph(style='Heading 2')
    set_para_format(p, before_pt=14, after_pt=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.text = ""
    r = p.add_run(text)
    set_run_font(r, "Times New Roman", 14, bold=True, color=COLOR_DARK_BLUE)
    return p

def add_h3(doc, text):
    p = doc.add_paragraph(style='Heading 3')
    set_para_format(p, before_pt=10, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.text = ""
    r = p.add_run(text)
    set_run_font(r, "Times New Roman", 13, bold=True, color=COLOR_MED_BLUE)
    return p

def add_centered_image(doc, img_path, width_inches=6.27):
    if not os.path.exists(img_path):
        print(f"WARNING: Image not found: {img_path}")
        return None
    p = doc.add_paragraph()
    set_para_format(p, before_pt=8, after_pt=4, line_spacing=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run()
    run.add_picture(img_path, width=Inches(width_inches))
    return p

def add_caption_xml(doc, label, chapter, seq_num, caption_text, is_first_in_chapter):
    r"""
    Inserts a caption paragraph with valid XML SEQ fields inside <w:r> tags.
    Format: "Figure {chapter}." + { SEQ Figure \r 1 \* ARABIC } + ": {caption_text}"
    """
    p = doc.add_paragraph()
    reset_clause = r' \r 1' if is_first_in_chapter else ''
    instr = f' SEQ {label}{reset_clause} \\* ARABIC '
    
    xml = f'''
    <w:p {nsdecls("w")}>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="120" w:after="240" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:iCs/>
          <w:sz w:val="22"/><w:szCs w:val="22"/>
        </w:rPr>
        <w:t>{label} {chapter}.</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:sz w:val="22"/>
        </w:rPr>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:sz w:val="22"/>
        </w:rPr>
        <w:instrText xml:space="preserve">{instr}</w:instrText>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:sz w:val="22"/>
        </w:rPr>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:noProof/><w:sz w:val="22"/>
        </w:rPr>
        <w:t>{seq_num}</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:sz w:val="22"/>
        </w:rPr>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
          <w:i/><w:sz w:val="22"/>
        </w:rPr>
        <w:t xml:space="preserve">: {caption_text}</w:t>
      </w:r>
    </w:p>
    '''
    elem = parse_xml(xml)
    p._p.getparent().replace(p._p, elem)
    return elem

def add_styled_table(doc, headers, data_rows, col_widths=None):
    """
    Creates a styled table adhering to golden specifications:
    - Table centered
    - Custom borders 0.5pt #1B3A5C
    - Header row: bold white on #1B3A5C background
    - Alternating row background for enhanced readability
    """
    table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Apply table borders via XML
    tblPr = table._tbl.tblPr
    borders_xml = f'''
    <w:tblBorders {nsdecls("w")}>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>
    </w:tblBorders>
    '''
    tblPr.append(parse_xml(borders_xml))

    # Format Header Row
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        set_para_format(p, before_pt=4, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(h_text)
        set_run_font(r, "Times New Roman", 11, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        shd_xml = f'<w:shd {nsdecls("w")} w:fill="{HEX_NAVY}" w:val="clear"/>'
        hdr_cells[i]._tc.get_or_add_tcPr().append(parse_xml(shd_xml))
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Format Data Rows
    for r_idx, row_data in enumerate(data_rows):
        row_cells = table.rows[r_idx + 1].cells
        bg_fill = HEX_LIGHT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = ""
            p = row_cells[c_idx].paragraphs[0]
            # Center shorter first column, justify/left align longer columns
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 or len(val) < 15 else WD_ALIGN_PARAGRAPH.LEFT
            set_para_format(p, before_pt=3, after_pt=3, line_spacing=1.15, align=align)
            r = p.add_run(val)
            set_run_font(r, "Times New Roman", 10.5, bold=False, color=RGBColor(0, 0, 0))
            if bg_fill != "FFFFFF":
                shd_xml = f'<w:shd {nsdecls("w")} w:fill="{bg_fill}" w:val="clear"/>'
                row_cells[c_idx]._tc.get_or_add_tcPr().append(parse_xml(shd_xml))
            row_cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Set column widths if provided
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)

    # Empty paragraph after table for spacing
    p_after = doc.add_paragraph()
    set_para_format(p_after, before_pt=0, after_pt=6, line_spacing=1.0)
    return table

def add_field_paragraph(doc, instr_text):
    """Inserts a field (like TOC) inside a paragraph with fldChar in <w:r>."""
    p = doc.add_paragraph()
    xml = f'''
    <w:p {nsdecls("w")}>
      <w:pPr>
        <w:spacing w:before="60" w:after="120" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve"> {instr_text} </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
    '''
    elem = parse_xml(xml)
    p._p.getparent().replace(p._p, elem)
    return elem

print("Helper definitions loaded successfully.")
