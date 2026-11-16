import docx
import zipfile
import os
import shutil
from docx.oxml import parse_xml
from report_builder_core import (
    TEMPLATE_DOCX, REPO_DIR, COLOR_DARK_BLUE, COLOR_MED_BLUE,
    set_run_font, set_para_format, add_paragraph_styled,
    add_h1, add_h2, add_h3, add_field_paragraph, UOH_LOGO
)

def make_clean_template():
    doc = docx.Document(TEMPLATE_DOCX)
    
    # 1. Capture Section 0 (Cover Page) XML properties
    s0_xml = doc.sections[0]._sectPr.xml

    # 2. Completely wipe ALL body elements (paragraphs, tables, orphan SDT / TOC)
    # preserving ONLY the final sectPr for Section 1
    body = doc._body._element
    for child in list(body):
        if child.tag.endswith('sectPr'):
            continue
        body.remove(child)

    # ---------------------------------------------------------
    # COVER PAGE (Section 0: No Header, No Footer, No Page Number)
    # ---------------------------------------------------------
    p_logo = doc.add_paragraph()
    set_para_format(p_logo, before_pt=20, after_pt=12, line_spacing=1.0, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r_logo = p_logo.add_run()
    if os.path.exists(UOH_LOGO):
        r_logo.add_picture(UOH_LOGO, height=docx.shared.Inches(1.8))

    p_uoh = doc.add_paragraph()
    set_para_format(p_uoh, before_pt=6, after_pt=4, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_uoh.add_run("University Of Hail")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_col = doc.add_paragraph()
    set_para_format(p_col, before_pt=2, after_pt=4, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_col.add_run("College of Computer Science and Engineering")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_dep = doc.add_paragraph()
    set_para_format(p_dep, before_pt=2, after_pt=36, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_dep.add_run("Department of [Department Name]")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_org_lbl = doc.add_paragraph()
    set_para_format(p_org_lbl, before_pt=18, after_pt=6, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_org_lbl.add_run("Training Organization Name :")
    set_run_font(r, "Times New Roman", 24, bold=True, color=COLOR_DARK_BLUE)

    p_org_name = doc.add_paragraph()
    set_para_format(p_org_name, before_pt=4, after_pt=40, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_org_name.add_run("[Training Organization Name]")
    set_run_font(r, "Times New Roman", 26, bold=True, color=COLOR_DARK_BLUE)

    p_adv_lbl = doc.add_paragraph()
    set_para_format(p_adv_lbl, before_pt=18, after_pt=4, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_adv_lbl.add_run("Under Supervision of Academic Advisor:")
    set_run_font(r, "Times New Roman", 18, bold=True, color=COLOR_DARK_BLUE)

    p_adv_name = doc.add_paragraph()
    set_para_format(p_adv_name, before_pt=2, after_pt=36, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_adv_name.add_run("Dr. [Advisor Name]")
    set_run_font(r, "Times New Roman", 20, bold=True, color=COLOR_DARK_BLUE)

    p_sub_lbl = doc.add_paragraph()
    set_para_format(p_sub_lbl, before_pt=18, after_pt=4, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_sub_lbl.add_run("Submitted by:")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_sub_val = doc.add_paragraph()
    set_para_format(p_sub_val, before_pt=2, after_pt=24, line_spacing=1.15, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    r = p_sub_val.add_run("[Student Name]        | [Student ID]")
    set_run_font(r, "Times New Roman", 18, bold=True, color=COLOR_DARK_BLUE)
    # Attach Section 0 properties to end Cover Page cleanly (no header, no footer, no page number)
    # This automatically transitions to Section 1 (Page 2) with headers and footers enabled
    p_sub_val._p.get_or_add_pPr().append(parse_xml(s0_xml))

    # ---------------------------------------------------------
    # SECTION 1: PRELIMINARY PAGES (Page Numbers, 100% English)
    # ---------------------------------------------------------
    # Abstract (Strict ~3/4 Page, ~220 words)
    add_h1(doc, "Abstract")
    add_paragraph_styled(doc, "[Insert executive abstract summarizing the training organization, core responsibilities, and capstone project here. Guideline: Exactly ~3/4 page in length, ~200-220 words across 2 structured paragraphs. Strictly 100% English with no Arabic characters or phrases.]")
    doc.add_page_break()

    # Acknowledgments (Strict <= 1/2 Page, ~110 words)
    add_h1(doc, "Acknowledgments")
    add_paragraph_styled(doc, "[Insert acknowledgments expressing gratitude to Almighty Allah, academic advisor (Dr. [Advisor Name]), hospital/organization leadership, faculty, and family here. Guideline: Half a page or less (<= 1/2 page, ~100-120 words). Strictly 100% English.]")
    doc.add_page_break()

    # Single Table of Contents (No duplicate before cover page)
    add_h1(doc, "Table of Contents")
    add_field_paragraph(doc, r'TOC \h \o "1-3"')
    doc.add_page_break()

    add_h1(doc, "Table of Figures")
    add_field_paragraph(doc, r'TOC \h \z \c "Figure"')
    doc.add_page_break()

    add_h1(doc, "Table of Tables")
    add_field_paragraph(doc, r'TOC \h \z \c "Table"')
    doc.add_page_break()

    add_h1(doc, "Introduction")
    add_paragraph_styled(doc, "[Insert preface/introduction explaining the cooperative training curriculum, academic context, and document structure.]")
    doc.add_page_break()

    # Chapter I
    add_h1(doc, "CHAPTER I:  INTRODUCTION")
    add_h2(doc, "1.1 Introduction")
    add_paragraph_styled(doc, "[Insert introductory context on the industry/sector and public health/regional transformation.]")
    add_h2(doc, "1.2 Training Organization")
    add_h3(doc, "Relevant Work")
    add_paragraph_styled(doc, "[Insert relevant institutional ecosystem and regional healthcare network overview.]")
    add_h3(doc, "1.2.1 What is Training Organization.?")
    add_paragraph_styled(doc, "[Insert introductory overview of the training entity. Note: Insert Figure 1.1 immediately beneath this opening paragraph. Use generic prose reference (e.g. 'As illustrated in the facility image...') without hardcoding static figure numbers.]")
    add_h3(doc, "1.2.2 History of Training Organization")
    add_paragraph_styled(doc, "[Insert historical background, founding milestones, and expansion phases.]")
    add_h3(doc, "1.2.3 Overview of Training Organization")
    add_paragraph_styled(doc, "[Insert overview of operational divisions, clinical departments, bed capacity, and specialized units.]")
    add_h3(doc, "Vision")
    add_paragraph_styled(doc, "[Insert organizational vision statement.]")
    add_h3(doc, "The Message")
    add_paragraph_styled(doc, "[Insert organizational mission statement.]")
    add_h3(doc, "The Objectives")
    add_paragraph_styled(doc, "[Insert institutional strategic objectives.]")
    doc.add_page_break()

    # Chapter II
    add_h1(doc, "CHAPTER II:  TASKS & EXPERIENCES")
    add_h2(doc, "2.1 Requirements Analysis")
    add_h3(doc, "2.1.1 Technical support")
    add_paragraph_styled(doc, "[Detail workstation maintenance, operating system deployment, hardware troubleshooting, and printer support. Note: Do NOT mention months or duration.]")
    add_h3(doc, "2.1.2 Apps and system")
    add_paragraph_styled(doc, "[Detail enterprise software deployment, security scans, data tracking in Excel, and user instructional manuals.]")
    add_h3(doc, "2.1.3 Networks & Communications")
    add_paragraph_styled(doc, "[Detail LAN auditing, IP address surveys, patch cabling, switch ports, and network printer sharing.]")
    add_h2(doc, "2.2 Experiences & Problems")
    add_h3(doc, "2.2.1 Experiences")
    add_paragraph_styled(doc, "[Detail rewarding milestones, technical challenges resolved, professional soft skills, and self-directed learning.]")
    doc.add_page_break()

    # Chapter III
    add_h1(doc, "CHAPTER III:  COOP PROJECT")
    add_h2(doc, "3.1 Problem Statement")
    add_paragraph_styled(doc, "[Define the domain challenges, inefficiencies, and motivation behind the software project.]")
    add_h2(doc, "3.2 Project Overview")
    add_paragraph_styled(doc, "[Provide a high-level architectural and functional summary of the developed application.]")
    add_h2(doc, "3.3 Project Objectives")
    add_paragraph_styled(doc, "[Articulate measurable project objectives alongside technical implementations.]")
    add_h2(doc, "3.4 Project Phases and Problem Solution")
    add_paragraph_styled(doc, "[Describe the systematic software engineering lifecycle phases from requirements to deployment.]")
    add_h2(doc, "3.5 Project Development Process")
    add_paragraph_styled(doc, "[Detail the technology stack, development tools, stakeholder roles, functional requirements, and non-functional guarantees.]")
    add_h2(doc, "3.6 System Design Diagrams")
    add_paragraph_styled(doc, "[Embed and analyze Use Case diagrams, Entity-Relationship diagrams, Sequence diagrams, and Activity flowcharts. Note: Refer to diagrams generically (e.g. 'As modeled in the Use Case diagram...', 'The sequence diagram demonstrates...') rather than static numbers.]")
    add_h2(doc, "3.7 System Interfaces Overview")
    add_paragraph_styled(doc, "[Present role-specific user interface screenshots, accompanied by at least two lines of functional explanation per image. Rule: Refer to screenshots generically without hardcoding static figure numbers in prose.]")
    add_h2(doc, "3.8 Testing")
    add_paragraph_styled(doc, "[Detail unit testing, integration testing, security vulnerability auditing, and user acceptance scenarios.]")
    doc.add_page_break()

    # Chapter IV
    add_h1(doc, "CHAPTER IV:  CONCLUSION & RECOMMENDATIONS")
    add_h2(doc, "Conclusion")
    add_paragraph_styled(doc, "[Synthesize key takeaways, professional growth, and capstone project achievements.]")
    add_h2(doc, "Recommendations")
    add_paragraph_styled(doc, "[Provide actionable recommendations for the training organization and academic institution.]")
    doc.add_page_break()

    # References
    add_h1(doc, "References")
    add_paragraph_styled(doc, "[Insert full bibliographic references in APA format with functional URLs. Strictly 100% English.]")

    clean_docx = os.path.join(REPO_DIR, 'template_term5.docx')
    doc.save(clean_docx)
    print(f'Saved clean template to {clean_docx}')

    # Unpack into Tamplate/ directory
    tamplate_dir = os.path.join(REPO_DIR, 'Tamplate')
    if os.path.exists(tamplate_dir):
        shutil.rmtree(tamplate_dir)
    os.makedirs(tamplate_dir, exist_ok=True)

    with zipfile.ZipFile(clean_docx, 'r') as z:
        z.extractall(tamplate_dir)
    print(f'Unpacked clean template into {tamplate_dir}')

if __name__ == '__main__':
    make_clean_template()
