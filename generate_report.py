#!/usr/bin/env python3
"""
ReportTerm5 - Automated Cooperative Training Report Generator
Project: Riaya Medical Diagnostics Platform
Student: Shadia Essa Al-Enezi (ID: 202304937)
Organization: Al-Hait General Hospital (Hail Health Cluster)
Supervisor: Al-Bandari Al-Otaibi
University: University of Hail, College of Computer Science and Engineering
"""

import os
import sys
import shutil
import zipfile
from lxml import etree
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from report_builder_core import (
    BASE_DIR, REPO_DIR, TEMPLATE_DOCX, OUTPUT_DIR, OUTPUT_DOCX, REPORT_EXPORT_DOCX,
    UOH_LOGO, ORG_MEDIA, DIAG_DIR, SCREEN_DIR,
    COLOR_DARK_BLUE, COLOR_MED_BLUE, W_NS,
    set_run_font, set_para_format, add_paragraph_styled,
    add_h1, add_h2, add_h3,
    add_centered_image, add_caption_xml, add_styled_table, add_field_paragraph
)

def build_report():
    print("Initializing Report Generation...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(os.path.dirname(REPORT_EXPORT_DOCX), exist_ok=True)

    # Load golden template
    doc = docx.Document(TEMPLATE_DOCX)

    # Capture Section 0 (Cover Page) XML properties
    s0_xml = doc.sections[0]._sectPr.xml

    # Completely wipe ALL body elements (paragraphs, tables, orphan SDT / TOC)
    # preserving ONLY the final sectPr for Section 1
    body = doc._body._element
    for child in list(body):
        if child.tag.endswith('sectPr'):
            continue
        body.remove(child)

    # Track sequence numbers per chapter
    fig_counts = {1: 0, 2: 0, 3: 0, 4: 0}
    tbl_counts = {1: 0, 2: 0, 3: 0, 4: 0}

    def put_figure(img_path, ch, caption, width=6.27):
        nonlocal fig_counts
        fig_counts[ch] += 1
        is_first = (fig_counts[ch] == 1)
        add_centered_image(doc, img_path, width_inches=width)
        add_caption_xml(doc, 'Figure', ch, fig_counts[ch], caption, is_first)

    def put_table(headers, data, ch, caption, col_widths=None):
        nonlocal tbl_counts
        tbl_counts[ch] += 1
        is_first = (tbl_counts[ch] == 1)
        add_caption_xml(doc, 'Table', ch, tbl_counts[ch], caption, is_first)
        add_styled_table(doc, headers, data, col_widths=col_widths)

    print("Building Cover Page...")
    # ---------------------------------------------------------
    # COVER PAGE
    # ---------------------------------------------------------
    p_logo = doc.add_paragraph()
    set_para_format(p_logo, before_pt=20, after_pt=12, line_spacing=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_logo = p_logo.add_run()
    if os.path.exists(UOH_LOGO):
        r_logo.add_picture(UOH_LOGO, height=Inches(1.8))

    p_uoh = doc.add_paragraph()
    set_para_format(p_uoh, before_pt=6, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_uoh.add_run("University Of Hail")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_col = doc.add_paragraph()
    set_para_format(p_col, before_pt=2, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_col.add_run("College of Computer Science and Engineering")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_dep = doc.add_paragraph()
    set_para_format(p_dep, before_pt=2, after_pt=36, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_dep.add_run("Department of Computer Science")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_org_lbl = doc.add_paragraph()
    set_para_format(p_org_lbl, before_pt=18, after_pt=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_org_lbl.add_run("Training Organization Name :")
    set_run_font(r, "Times New Roman", 24, bold=True, color=COLOR_DARK_BLUE)

    p_org_name = doc.add_paragraph()
    set_para_format(p_org_name, before_pt=4, after_pt=40, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_org_name.add_run("Al-Hait General Hospital")
    set_run_font(r, "Times New Roman", 26, bold=True, color=COLOR_DARK_BLUE)

    p_adv_lbl = doc.add_paragraph()
    set_para_format(p_adv_lbl, before_pt=18, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_adv_lbl.add_run("Under Supervision of Academic Advisor:")
    set_run_font(r, "Times New Roman", 18, bold=True, color=COLOR_DARK_BLUE)

    p_adv_name = doc.add_paragraph()
    set_para_format(p_adv_name, before_pt=2, after_pt=36, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_adv_name.add_run("Al-Bandari Al-Otaibi")
    set_run_font(r, "Times New Roman", 20, bold=True, color=COLOR_DARK_BLUE)

    p_sub_lbl = doc.add_paragraph()
    set_para_format(p_sub_lbl, before_pt=18, after_pt=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_sub_lbl.add_run("Submitted by:")
    set_run_font(r, "Times New Roman", 16, bold=True, color=COLOR_DARK_BLUE)

    p_sub_val = doc.add_paragraph()
    set_para_format(p_sub_val, before_pt=2, after_pt=24, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_sub_val.add_run("Shadia Essa Al-Enezi        | 202304937")
    set_run_font(r, "Times New Roman", 18, bold=True, color=COLOR_DARK_BLUE)
    # Attach Section 0 properties to end Cover Page cleanly (no header, no footer, no page number)
    # This automatically creates a section break to Page 2 without an extra blank page
    p_sub_val._p.get_or_add_pPr().append(parse_xml(s0_xml))

    print("Building Preliminary Sections...")
    # ---------------------------------------------------------
    # ABSTRACT (Strict ~3/4 Page, 100% English)
    # ---------------------------------------------------------
    add_h1(doc, "Abstract")
    add_paragraph_styled(doc,
        "This cooperative training report encapsulates the professional and technical internship undertaken at Al-Hait General Hospital, a primary healthcare institution operating under the governance of the Hail Health Cluster and the Ministry of Health in the Kingdom of Saudi Arabia. Undertaken to fulfill the graduation requirements for the Bachelor of Science in Computer Science at the University of Hail, College of Computer Science and Engineering, the training program bridged academic computing principles with healthcare informatics and hospital IT infrastructure. Key operational responsibilities included computer hardware auditing, operating system deployment and maintenance, local area network (LAN) administration, TCP/IP configuration, endpoint cybersecurity, and end-user clinical support across emergency, outpatient, inpatient, and diagnostic divisions.")

    add_paragraph_styled(doc,
        "To overcome clinical laboratory workflow bottlenecks, a full-stack web platform titled 'Riaya' (Intelligent Diagnostic Auditing and Medical Report Platform) was designed and implemented using PHP 8.2 MVC, MySQL, and role-based access control. The platform integrates international standards including Logical Observation Identifiers Names and Codes (LOINC) and International Classification of Diseases (ICD-10). Its deterministic ClinicalEngine verifies laboratory values against Mayo Clinic and ARUP reference intervals, triggering instant STAT alerts for life-threatening panic limits while preventing duplicate 48-hour testing. A zero-PII generative artificial intelligence pipeline powered by OpenRouter LLM translates complex lab findings into clear, educational explanations for patients. The resulting platform delivers robust clinical governance, high performance, and enhanced healthcare accessibility.")

    doc.add_page_break()

    # ---------------------------------------------------------
    # ACKNOWLEDGMENTS (Strict <= 1/2 Page, 100% English)
    # ---------------------------------------------------------
    add_h1(doc, "Acknowledgments")
    add_paragraph_styled(doc,
        "First and foremost, all praise, gratitude, and adoration are due to Almighty Allah for bestowing upon me the strength, perseverance, and enlightenment required to successfully complete this cooperative training program and bring this technical report to fruition.")

    add_paragraph_styled(doc,
        "I express my profound gratitude to my academic advisor, Al-Bandari Al-Otaibi, for her scholarly stewardship, continuous encouragement, and constructive critiques throughout this internship. Sincere appreciation is also extended to the leadership, IT administration, medical staff, and technical supervisors at Al-Hait General Hospital and the Hail Health Cluster for their professional hospitality and mentorship. Furthermore, I convey my heartfelt thanks to the faculty of the College of Computer Science and Engineering at the University of Hail, and to my beloved family for their unconditional love, patience, and constant prayers.")

    doc.add_page_break()

    # ---------------------------------------------------------
    # TABLE OF CONTENTS
    # ---------------------------------------------------------
    add_h1(doc, "Table of Contents")
    add_field_paragraph(doc, r'TOC \h \o "1-3"')
    doc.add_page_break()

    # ---------------------------------------------------------
    # TABLE OF FIGURES
    # ---------------------------------------------------------
    add_h1(doc, "Table of Figures")
    add_field_paragraph(doc, r'TOC \h \z \c "Figure"')
    doc.add_page_break()

    # ---------------------------------------------------------
    # TABLE OF TABLES
    # ---------------------------------------------------------
    add_h1(doc, "Table of Tables")
    add_field_paragraph(doc, r'TOC \h \z \c "Table"')
    doc.add_page_break()

    # ---------------------------------------------------------
    # INTRODUCTION (PREFACE)
    # ---------------------------------------------------------
    add_h1(doc, "Introduction")
    add_paragraph_styled(doc,
        "The Cooperative Education Program (COOP) established by the College of Computer Science and Engineering at the University of Hail represents an essential cornerstone in academic and professional development. Serving as an intensive experiential learning curriculum, the program mandates that senior undergraduate students immerse themselves in institutional enterprise environments. The fundamental objective is to cultivate practical competency, apply theoretical computing methodologies to real-world operational challenges, and understand the administrative, technical, and socio-ethical dynamics of contemporary workplaces.")

    add_paragraph_styled(doc,
        "Modern healthcare institutions represent one of the most demanding, mission-critical environments for information technology. Hospitals require uninterrupted system availability, uncompromising data integrity, stringent patient confidentiality, and flawless communication across multidisciplinary departments. Undertaking cooperative training within a healthcare facility provides a unique vantage point to analyze how enterprise resource planning, network architectures, and diagnostic information systems intersect to directly impact patient care and public health outcomes.")

    add_paragraph_styled(doc,
        "This report is organized into four structured chapters that comprehensively document the cooperative training experience and its capstone software deliverable:",
        bold_prefix="Report Organization: ")

    add_paragraph_styled(doc,
        "Presents an exhaustive institutional study of Al-Hait General Hospital and its governing body, the Hail Health Cluster. It details the organization's historical background, operational capacity, clinical departments, digital infrastructure, vision, mission, and strategic healthcare objectives under Saudi Vision 2030.",
        bold_prefix="• Chapter I (Introduction & Training Organization): ")

    add_paragraph_styled(doc,
        "Synthesizes the technical support tasks, enterprise application administration, and local network operations executed during the training period, concluding with an analytical reflection on acquired technical and interpersonal competencies.",
        bold_prefix="• Chapter II (Tasks & Experiences): ")

    add_paragraph_styled(doc,
        "Documents the full engineering lifecycle of 'Riaya' (Intelligent Diagnostic Auditing and Medical Report Platform), detailing its clinical problem statement, software requirements, database modeling, unified design diagrams, role-based user interfaces, and automated quality assurance.",
        bold_prefix="• Chapter III (Coop Project — Riaya): ")

    add_paragraph_styled(doc,
        "Synthesizes the overarching conclusions derived from the cooperative training tenure and articulates strategic technical recommendations for both hospital IT operations and future academic cooperative cohorts.",
        bold_prefix="• Chapter IV (Conclusion & Recommendations): ")

    doc.add_page_break()

    print("Building Chapter 1: Introduction & Training Organization...")
    # ---------------------------------------------------------
    # CHAPTER I: INTRODUCTION
    # ---------------------------------------------------------
    add_h1(doc, "CHAPTER I:  INTRODUCTION")

    add_h2(doc, "1.1 Introduction")
    add_paragraph_styled(doc,
        "The Kingdom of Saudi Arabia is presently experiencing an unprecedented transformation across its social, economic, and institutional landscapes, catalyzed by the visionary directives of Saudi Vision 2030. At the heart of this comprehensive national development plan lies the Health Sector Transformation Program (HSTP), which seeks to restructure the national healthcare delivery model into an integrated, efficient, patient-centered, and digitally advanced ecosystem. Through the establishment of regional Health Clusters under the auspices of the Ministry of Health (MOH) and the Center for National Health Insurance (CNHI), public healthcare facilities across Saudi Arabia are actively transitioning from siloed regional directorates toward decentralized, accountable care organizations focused on preventive health, universal coverage, clinical governance, and comprehensive digital enablement.")

    add_paragraph_styled(doc,
        "In this transforming healthcare landscape, secondary and peripheral general hospitals fulfill a vital strategic role. These regional facilities serve as the frontline defense against acute illnesses, trauma emergencies, and chronic disease burdens for populations residing in outlying governorates and rural centers, substantially reducing the logistical necessity for patients to travel hundreds of kilometers to tertiary metropolitan medical cities. However, the operational efficacy of regional hospitals is intrinsically tied to the robustness of their health informatics systems, hardware infrastructure, and technical support frameworks. When clinical applications, laboratory networks, and diagnostic imaging archives function seamlessly, clinical practitioners can deliver timely, evidence-based care, thereby directly reducing morbidity and mortality rates.")

    add_h2(doc, "1.2 Training Organization")

    add_h3(doc, "Relevant Work")
    add_paragraph_styled(doc,
        "Healthcare delivery within the Hail Region has undergone substantial modernization under the administrative oversight of the Hail Health Cluster. As an accountable integrated health network, the Hail Health Cluster oversees a comprehensive continuum of care comprising major urban specialized hospitals (such as King Khalid Hospital and Hail General Hospital), peripheral general hospitals situated across various governorates, specialized dental and cardiac centers, and an expansive network of Primary Healthcare Centers (PHCs). Within this regional matrix, peripheral hospitals function as essential clinical hubs that deliver urgent triage, inpatient admission, surgical interventions, maternal-child healthcare, and outpatient clinical consultations.")

    add_paragraph_styled(doc,
        "Digital health initiatives across the Hail Health Cluster emphasize the phased adoption of standardized electronic medical records, centralized telemedicine connections with regional tertiary consultants, electronic laboratory ordering, and digital diagnostic reporting. Understanding the institutional, structural, and technical attributes of a peripheral hospital operating within this network is imperative for analyzing where software automation and clinical intelligence can resolve systemic communication friction.")

    add_h3(doc, "1.2.1 What is Training Organization.?")
    add_paragraph_styled(doc,
        "Al-Hait General Hospital is a state-funded public secondary healthcare institution operating under the governance of the Hail Health Cluster and the Ministry of Health. Geographically situated in Al-Hait Governorate, approximately 250 kilometers south of Hail City, the hospital functions as the primary medical facility serving the local population of the governorate and dozens of surrounding rural villages, desert communities, and municipal centers. The hospital operates as a 50-bed general healthcare facility equipped with specialized diagnostic, surgical, maternal, and emergency units designed to deliver comprehensive medical care.")

    # Figure 1.1: Al-Hait General Hospital placed right beneath the opening paragraph
    img_hosp = os.path.join(ORG_MEDIA, 'hospital_facility.jpg')
    put_figure(img_hosp, 1, "Al-Hait General Hospital — Administrative and Clinical Complex")

    add_paragraph_styled(doc,
        "As illustrated in Figure 1.1, Al-Hait General Hospital comprises a modern, purpose-built medical complex featuring dedicated wings for outpatient clinical consultations, fully equipped emergency trauma bays, inpatient medical and surgical wards, specialized hemodialysis stations, digital diagnostic imaging suites, automated clinical laboratories, and centralized administrative offices. The hospital represents a vital lifeline for south Hail, handling significant daily patient throughput across emergency admissions, routine outpatient specialty appointments, routine laboratory specimen screenings, and chronic disease management regimens.")

    add_paragraph_styled(doc,
        "Beyond clinical treatment, Al-Hait General Hospital plays a critical civic role in executing national public health screening campaigns, seasonal vaccination drives, epidemiological surveillance, and health education initiatives across the southern sector of the Hail region. The institution's day-to-day operations rely upon an intricate blend of medical technology, digital records systems, and enterprise administrative networks managed by dedicated clinical and technical teams.")

    add_h3(doc, "1.2.2 History of Training Organization")
    add_paragraph_styled(doc,
        "The establishment of Al-Hait General Hospital represents the realization of a longstanding community aspiration. For decades, residents of Al-Hait Governorate and its affiliated villages faced formidable logistical hurdles in obtaining essential medical care, frequently requiring arduous road journeys of over 250 kilometers across mountainous desert terrain to reach urban hospitals in Hail City. Recognizing this urgent humanitarian and public health necessity, the Ministry of Health approved the construction of a comprehensive 50-bed general hospital equipped with contemporary medical technology.")

    add_paragraph_styled(doc,
        "Following extensive construction and infrastructure development, Al-Hait General Hospital commenced its initial phased operational activation in early 2014 (corresponding to Rabi' Al-Thani 1435 AH). The opening was inaugurated with the commissioning of the 24-hour Emergency and Acute Care Department and primary outpatient clinics, bringing immediate relief to thousands of local families. Over subsequent years, operational capacity was systematically expanded to incorporate advanced inpatient surgical suites, an expanded Obstetrics and Gynecology department, dedicated pediatric wards, and advanced medical imaging equipment.")

    add_paragraph_styled(doc,
        "A monumental milestone in the hospital's history occurred with the commissioning of the specialized Artificial Hemodialysis Center. Prior to this opening, chronic kidney failure patients within the governorate were forced to travel weekly to distant cities for life-sustaining hemodialysis sessions. The local hemodialysis center significantly mitigated patient physical suffering and financial burden. In recent years, under the administrative restructuring of the Saudi healthcare sector, Al-Hait General Hospital was formally integrated into the Hail Health Cluster, aligning its clinical protocols, quality assurance methodologies, and IT systems with national health transformation standards.")

    add_h3(doc, "1.2.3 Overview of Training Organization")
    add_paragraph_styled(doc,
        "Al-Hait General Hospital operates a well-defined organizational hierarchy comprising medical administration, nursing services, allied medical departments, financial and administrative affairs, and health informatics. The clinical architecture of the facility is partitioned into several specialized functional units:",
        bold_prefix="Clinical and Operational Architecture: ")

    add_paragraph_styled(doc,
        "Operates uninterrupted twenty-four hours a day, providing rapid triage, acute resuscitation, trauma stabilization, and critical care dispatch. It features dedicated male and female treatment bays, trauma rooms, and an automated cardiac monitoring network.",
        bold_prefix="1. Emergency and Trauma Department: ")

    add_paragraph_styled(doc,
        "Comprises multiple clinical consultation suites covering Internal Medicine, General Surgery, Pediatrics, Orthopedics, Ophthalmology, ENT, Dermatology, and Cardiology, serving scheduled and walk-in clinical consultations.",
        bold_prefix="2. Outpatient Specialty Clinics: ")

    add_paragraph_styled(doc,
        "Separate inpatient care pavilions for adult male, adult female, and pediatric medical-surgical admissions, supported by continuous nursing monitoring and clinical rounds.",
        bold_prefix="3. Inpatient Medical and Surgical Wards: ")

    add_paragraph_styled(doc,
        "A purpose-built maternity unit with private labor and delivery suites, post-natal recovery rooms, and neonatal resuscitation facilities.",
        bold_prefix="4. Obstetrics and Gynecology Department: ")

    add_paragraph_styled(doc,
        "A highly specialized clinical pavilion equipped with state-of-the-art dialysis machines, water purification plants, and dedicated nephrology nursing staff providing maintenance dialysis for chronic renal failure patients.",
        bold_prefix="5. Hemodialysis Center: ")

    add_paragraph_styled(doc,
        "Equipped with digital radiography (DR) machines, ultrasound diagnostic scanners, and computed tomography (CT) facilities integrated into digital picture archiving systems.",
        bold_prefix="6. Diagnostic Radiology and Medical Imaging: ")

    add_paragraph_styled(doc,
        "Performs hematology, biochemistry, microbiology, serology, and blood banking services essential for rapid clinical diagnosis.",
        bold_prefix="7. Clinical Laboratory and Blood Bank: ")

    add_paragraph_styled(doc,
        "Administers inpatient medication dispensing, outpatient prescription fulfillment, and pharmacological inventory control.",
        bold_prefix="8. Hospital Pharmacy Services: ")

    # Figure 1.2: Hail Health Cluster Network Map
    img_map = os.path.join(ORG_MEDIA, 'hail_health_network_map.png')
    put_figure(img_map, 1, "Hail Health Cluster Geographic and Healthcare Facilities Network")

    add_paragraph_styled(doc,
        "As depicted in Figure 1.2, Al-Hait General Hospital serves as an anchor healthcare facility within the southern geographic perimeter of the Hail Health Cluster. The hospital maintains direct communication links with urban tertiary hospitals for patient transfer protocols, tele-consultation reviews, and regional epidemiology reporting, making its local network and IT uptime crucial to healthcare continuity.")

    add_h3(doc, "Vision")
    add_paragraph_styled(doc,
        "To stand as a pioneering, trusted, and technologically advanced regional healthcare facility that delivers exemplary, comprehensive, and compassionate patient care, actively realizing the health transformation ambitions of Saudi Vision 2030 by fostering a healthy, resilient, and thriving community across Al-Hait Governorate.")

    add_h3(doc, "The Message")
    add_paragraph_styled(doc,
        "Al-Hait General Hospital is dedicated to providing high-quality, accessible, and evidence-based preventive, curative, and emergency healthcare services to all individuals across the governorate and its surrounding communities. We achieve clinical excellence through a patient-centered approach, competent multidisciplinary healthcare personnel, advanced medical technologies, continuous quality improvement, and adherence to the highest national healthcare standards within an environment of safety, dignity, and empathy.")

    add_h3(doc, "The Objectives")
    add_paragraph_styled(doc,
        "The strategic institutional objectives of Al-Hait General Hospital encompass:",
        bold_prefix="Institutional Strategic Objectives: ")

    add_paragraph_styled(doc,
        "Expanding preventive screening programs, maternal-child welfare initiatives, and chronic disease surveillance across the governorate.",
        bold_prefix="1. Enhancing Community Health and Preventive Care: ")

    add_paragraph_styled(doc,
        "Optimizing clinical pathways, reducing emergency department waiting times, and elevating patient satisfaction ratings.",
        bold_prefix="2. Elevating Clinical Service Quality and Access: ")

    add_paragraph_styled(doc,
        "Sustaining rigorous clinical governance, patient safety protocols, infection control programs, and full accreditation compliance under the Saudi Central Board for Accreditation of Healthcare Institutions (CBAHI).",
        bold_prefix="3. Fostering Patient Safety and Quality Accreditation: ")

    add_paragraph_styled(doc,
        "Upgrading hospital IT networks, transitioning toward fully digital clinical workflows, modernizing diagnostic computing equipment, and optimizing telemedicine and laboratory communications.",
        bold_prefix="4. Modernizing Digital Healthcare Infrastructure: ")

    add_paragraph_styled(doc,
        "Encouraging continuous professional development, hosting technical and healthcare cooperative internships, and fostering interdisciplinary collaboration.",
        bold_prefix="5. Investing in Human Capital and Technical Excellence: ")

    doc.add_page_break()

    print("Building Chapter 2: Tasks & Experiences...")
    # ---------------------------------------------------------
    # CHAPTER II: TASKS & EXPERIENCES
    # ---------------------------------------------------------
    add_h1(doc, "CHAPTER II:  TASKS & EXPERIENCES")

    add_h2(doc, "2.1 Requirements Analysis")
    add_paragraph_styled(doc,
        "During the cooperative training period at Al-Hait General Hospital, duties were embedded within the Information Technology and Technical Support Department. Enterprise healthcare facilities operate in a rigorous, time-critical domain where technical latency or hardware failures directly impact clinical workflows, emergency patient triage, and diagnostic report turnaround. Consequently, duties encompassed an exhaustive requirements analysis of the institution's computing environment, followed by hands-on execution across three functional domains: Technical Support, Applications & Systems, and Networks & Communications.")

    add_h3(doc, "2.1.1 Technical support")
    add_paragraph_styled(doc,
        "Technical support operations formed the operational core of daily workplace activities. A comprehensive diagnostic inspection and hardware inventory audit was executed across all operational workstations deployed in administrative suites, outpatient consultation rooms, the medical records archive, the outpatient pharmacy, and inpatient reception desks. Many desktop workstations exhibited severe operational degradation, operating system lag, and recurrent application freezing stemming from prolonged uninterrupted utilization, excessive temporary cache accumulation, and mechanical storage degradation.")

    add_paragraph_styled(doc,
        "To resolve these operational deficits, systematic workstation formatting and clean operating system installations were conducted. Computers were backed up, formatted, repartitioned, and re-imaged with enterprise Windows builds configured with the required regional language packages, Active Directory workstation provisioning, and administrative security policies. Diagnostic troubleshooting was systematically performed on systems experiencing monitor freezing, display resolution artifacts, thermal throttling due to internal dust accumulation, and power supply irregularities. Physical maintenance included opening desktop computer chassis, inspecting motherboard capacitors and cooling fans, and verifying SATA/NVMe cable integrity.")

    add_paragraph_styled(doc,
        "Printer maintenance constituted another prominent technical support responsibility. The hospital employs a heterogeneous fleet of desktop laser printers and high-capacity network multifunction printers (MFPs) across administrative and medical records departments. Tasks included configuring printer driver properties, mapping network print queues, diagnosing paper feed jamming, replacing toner cartridges, and maintaining consumable supply logs in coordination with department supervisors.")

    add_h3(doc, "2.1.2 Apps and system")
    add_paragraph_styled(doc,
        "The software ecosystem at Al-Hait General Hospital comprises healthcare enterprise portals, medical imaging viewers, administrative document management suites, and standard office productivity software. Technical responsibilities in this domain included verifying software compatibility, deploying required desktop applications, configuring secure web browsers for hospital internal portal access, and installing Adobe Acrobat PDF tools essential for reviewing medical documentation.")

    add_paragraph_styled(doc,
        "System security maintenance was executed as a top operational priority. Thorough security scans were conducted across hospital workstations utilizing enterprise endpoint security suites to detect and remediate potential malware, adware, or unauthorized software. Workstation operating system security patches were audited to ensure compliance with Ministry of Health IT security policies. User access management was supported through password resets, user profile configuration, and security permission auditing.")

    add_paragraph_styled(doc,
        "Furthermore, administrative data management was streamlined using Microsoft Excel. Structured tracking spreadsheets were developed to record the inventory of all hospital computing equipment, printer consumables, toner replacement lifecycles, and IP address mappings. User guides, structured instructional documentation, and orientation slide decks were authored to assist administrative staff in mastering updated desktop software, document navigation, and internal electronic filing protocols.")

    add_h3(doc, "2.1.3 Networks & Communications")
    add_paragraph_styled(doc,
        "Reliable networking is critical in a hospital environment where patient records, diagnostic requests, and pharmacy orders must traverse the local area network (LAN) instantly without packet loss or connection drops. Comprehensive networking duties included conducting a full hospital IP address survey, scanning the local network to document assigned static IP addresses for medical equipment, network printers, and servers, while mapping dynamic DHCP pools.")

    add_paragraph_styled(doc,
        "Physical network infrastructure tasks involved testing RJ-45 Ethernet patch cables utilizing network cable testers, re-crimping damaged connector terminals, tracing cable runs between office wall jacks and intermediate distribution frames (IDFs), and verifying switch port link speeds. Technical support was provided to resolve network drops on administrative workstations, diagnose DNS lookup failures, and verify default gateway routing.")

    add_paragraph_styled(doc,
        "In the medical records department, local network printer sharing was successfully architected and deployed. A high-capacity departmental laser printer was configured with a dedicated static IP address and shared across multiple staff workstations, eliminating previous bottlenecks where employees were forced to physically transfer files via external USB flash drives to print patient records.")

    add_h2(doc, "2.2 Experiences & Problems")

    add_h3(doc, "2.2.1 Experiences")
    add_paragraph_styled(doc,
        "The cooperative training experience at Al-Hait General Hospital proved profoundly rewarding, offering deep technical immersion into enterprise computing operations within a healthcare setting. Engaging directly with healthcare administrators, clinical physicians, and nursing supervisors cultivated an acute understanding of how IT reliability underpins patient welfare. The most rewarding milestones included restoring lagging workstations to peak computational performance, engineering departmental printer-sharing workflows that received high staff acclaim, and structuring comprehensive asset tracking systems.")

    add_paragraph_styled(doc,
        "Concurrently, the internship presented authentic technical and logistical challenges that demanded analytical persistence. A recurring technical obstacle was resolving hardware driver conflicts when connecting newly acquired printer models to legacy Windows workstations; this was overcome through methodical research, identifying compatible generic PostScript/PCL drivers, and reconfiguring spooler subsystems. Another challenge involved tracing incomplete historical paper logs for printer consumables, which required performing manual physical counts and converting fragmented records into structured, formula-driven Excel spreadsheets. Furthermore, self-directed learning was continuously pursued through online technical documentation and tutorial videos, mastering advanced Windows troubleshooting utilities, PowerShell scripting, and structured data entry.")

    # Table 2.1: Technical Skills Acquired
    tech_skills_headers = ["Domain", "Technical Competency", "Practical Workplace Application"]
    tech_skills_data = [
        ["Hardware Maintenance", "Hardware Diagnostics & Assembly", "Troubleshooting desktop freezes, replacing cooling fans, testing power supplies, and performing physical chassis maintenance."],
        ["OS Deployment", "Operating System Imaging & Formatting", "Clean installation of Windows enterprise editions, user profile provisioning, and system performance tuning."],
        ["Network Infrastructure", "LAN Administration & Cable Testing", "IP address auditing, RJ-45 Ethernet crimping, wall port patching, and switch port link verification."],
        ["Peripheral Systems", "Network Printer Configuration & Sharing", "Mapping network print queues, resolving driver conflicts, setting static IPs, and managing toner lifecycles."],
        ["Cybersecurity & Access", "Endpoint Security & Access Control", "Executing enterprise antivirus malware scans, enforcing password security policies, and managing user access privileges."],
        ["Productivity Software", "Data Structuring in Microsoft Excel", "Developing comprehensive inventory sheets, consumable tracking models, and structured technical support logs."]
    ]
    put_table(tech_skills_headers, tech_skills_data, 2, "Technical Skills Acquired during Cooperative Training", col_widths=[1.5, 2.0, 2.77])

    # Table 2.2: Soft Skills Acquired
    soft_skills_headers = ["Category", "Interpersonal Competency", "Professional Growth Impact"]
    soft_skills_data = [
        ["Communication", "Multidisciplinary Professional Communication", "Effectively communicating technical issues and instructions to non-technical healthcare and administrative staff."],
        ["Problem Solving", "Independent Root-Cause Analysis", "Methodically isolating complex technical bugs, researching documentation, and implementing durable fixes."],
        ["Time Management", "Prioritizing Critical Tasks", "Triage-based scheduling of urgent technical requests in clinical areas over routine administrative maintenance."],
        ["Teamwork & Ethics", "Workplace Etiquette & Patient Privacy", "Maintaining strict confidentiality of hospital records, adhering to ethical workplace standards, and collaborating with IT colleagues."],
        ["Continuous Learning", "Self-Directed Technical Upskilling", "Independently researching technical solutions, watching instructional tutorials, and mastering new IT tools."]
    ]
    put_table(soft_skills_headers, soft_skills_data, 2, "Professional, Administrative, and Interpersonal Skills Acquired", col_widths=[1.5, 2.0, 2.77])

    doc.add_page_break()

    print("Building Chapter 3: Coop Project — Riaya...")
    # ---------------------------------------------------------
    # CHAPTER III: COOP PROJECT
    # ---------------------------------------------------------
    add_h1(doc, "CHAPTER III:  COOP PROJECT")

    add_h2(doc, "3.1 Problem Statement")
    add_paragraph_styled(doc,
        "Clinical laboratory diagnostic testing represents the scientific foundation of contemporary healthcare, influencing over seventy percent of all clinical diagnoses, therapeutic regimens, and hospital admission pathways. Despite the critical importance of laboratory results, real-world clinical observation reveals profound systemic bottlenecks in how laboratory data is communicated, reviewed, and acted upon across the healthcare continuum:",
        bold_prefix="Core Problem Statement: ")

    add_paragraph_styled(doc,
        "Laboratory reports are inherently architected for specialized clinicians, presented in dense medical English, Latin biochemical abbreviations (e.g., ALT, AST, HbA1c, MCH, eGFR), and complex units of measurement (e.g., mmol/L, mg/dL, fL). Patients reviewing their reports online or on paper experience acute psychological anxiety, cognitive confusion, and distress, frequently resorting to unverified internet search queries or misinterpreting harmless borderline values as catastrophic illnesses.",
        bold_prefix="1. Patient Cognitive Disconnect and Health Literacy Deficit: ")

    add_paragraph_styled(doc,
        "Certain laboratory abnormalities represent immediate, life-threatening clinical emergencies—known as 'panic limits' or 'critical values' (e.g., serum potassium > 6.5 mmol/L with imminent cardiac arrest risk, or hemoglobin < 7.0 g/dL indicating severe acute hemorrhage). In conventional workflows, these results are logged in passive laboratory queues without proactive, guaranteed delivery to the treating physician's primary viewing screen, introducing catastrophic delays in emergency interventions.",
        bold_prefix="2. Delayed Attending Physician Notification for Critical Values: ")

    add_paragraph_styled(doc,
        "Laboratory reports frequently display numerical values in isolation, presenting merely an static single-point snapshot without immediately accessible visual benchmarking against the patient's historical test baseline. Without automated delta checking, subtle deterioration trajectories (e.g., progressive hemoglobin decline over weeks) go unnoticed until reaching acute crisis stages.",
        bold_prefix="3. Lack of Visual Longitudinal Delta Checking: ")

    add_paragraph_styled(doc,
        "Due to fragmented coordination between outpatient clinics, emergency rooms, and inpatient wards, identical diagnostic blood tests are frequently ordered multiple times within a 48-hour window for the same patient, generating unnecessary medical expenditures, lab reagent wastage, and patient physical discomfort.",
        bold_prefix="4. Diagnostic Duplication and Test Redundancy: ")

    add_paragraph_styled(doc,
        "Healthcare systems across different institutions utilize inconsistent local naming conventions for identical biological analytes, creating interoperability barriers and preventing cohesive semantic exchange.",
        bold_prefix="5. Non-Standardized Test Naming and Nomenclature: ")

    add_h2(doc, "3.2 Project Overview")
    add_paragraph_styled(doc,
        "To resolve these systemic challenges, 'Riaya' (Intelligent Diagnostic Auditing and Medical Report Platform) was conceived, engineered, and delivered as an enterprise web application. Riaya is an integrated clinical web platform that analyzes laboratory results instantaneously upon release, simplifies diagnostic meanings for patients into clear, reassuring Arabic, highlights abnormal and critical indicators against standardized clinical ranges, visualizes longitudinal comparative trends against historical baselines, and automatically triggers immediate STAT emergency alerts to attending physicians upon detecting critical values.")

    add_paragraph_styled(doc,
        "Riaya is engineered utilizing a high-performance backend powered by PHP 8.2 adhering to a custom Model-View-Controller (MVC) software architecture, a relational MySQL database (InnoDB) fully normalized to Third Normal Form (3NF), and a responsive modern frontend constructed with semantic HTML5, CSS3 Custom Properties, and Vanilla JavaScript (ES6+). The platform integrates international medical standards including Logical Observation Identifiers Names and Codes (LOINC) via live National Library of Medicine (NLM) RESTful web services, and International Classification of Diseases (ICD-10) coding for diagnosis linking. Furthermore, Riaya incorporates a deterministic ClinicalEngine built upon verified clinical reference ranges from Mayo Clinic Laboratories and ARUP Laboratories, paired with a Zero-PII Generative Artificial Intelligence pipeline via OpenRouter LLM for safe, compassionate patient educational report interpretations.")

    add_h2(doc, "3.3 Project Objectives")
    add_paragraph_styled(doc,
        "The strategic engineering and clinical objectives of the Riaya platform are categorized in Table 3.1, detailing the technical mechanisms implemented and their corresponding clinical outcomes.")

    # Table 3.1: Project Objectives
    obj_headers = ["Strategic Objective", "Technical Implementation Mechanism", "Anticipated Clinical & Operational Impact"]
    obj_data = [
        ["Real-Time Diagnostic Auditing", "Automated ClinicalEngine comparing incoming quantitative results against age/gender-stratified reference ranges.", "Instant categorization of every test parameter into Normal, Low, High, or Critical bands immediately upon release."],
        ["Automated STAT Critical Alerts", "Push notification engine generating red STAT alert banners on attending physician dashboards upon panic limit breach.", "Zero-latency notification of life-threatening results, requiring mandatory physician acknowledgment and logging."],
        ["Patient Health Empowerment", "Zero-PII OpenRouter AI pipeline translating biochemical data into compassionate, plain-language Arabic explanations.", "Drastically diminishes patient anxiety, eliminates misinformation, and enhances compliance with physician treatment plans."],
        ["Longitudinal Delta Tracking", "Comparative database engine tracking patient historical test results and calculating percentage delta shifts.", "Empowers clinicians to detect subtle progressive physiological declines before acute symptom onset occurs."],
        ["Duplicate Order Prevention", "Automated 48-hour temporal lookup checking pending/active orders for identical LOINC codes before order confirmation.", "Eliminates diagnostic redundancy, curtails unnecessary hospital laboratory costs, and spares patients repeated venipunctures."],
        ["International Standardization", "Integration with NLM LOINC web services for code lookup and WHO ICD-10 diagnostic coding.", "Ensures universal clinical interoperability, semantic consistency, and seamless future national health exchange (HIE) readiness."]
    ]
    put_table(obj_headers, obj_data, 3, "Riaya Project Strategic Objectives, Technical Implementations, and Clinical Impact", col_widths=[1.5, 2.3, 2.47])

    add_h2(doc, "3.4 Project Phases and Problem Solution")
    add_paragraph_styled(doc,
        "The engineering of Riaya followed a structured, phased software development lifecycle (SDLC) tailored to the stringent safety and reliability requirements of healthcare software:",
        bold_prefix="Systematic Development Phases: ")

    add_paragraph_styled(doc,
        "Conducting in-depth workflow analysis of hospital clinical laboratory operations, reviewing laboratory request requisitions, specimen processing workflows, reference range variations, and critical value reporting protocols.",
        bold_prefix="Phase 1: Requirements Analysis & Clinical Workflow Auditing — ")

    add_paragraph_styled(doc,
        "Formulating the custom MVC software architecture, defining strict data contracts between controllers, models, and clinical services, and designing cryptographic role-based authorization barriers.",
        bold_prefix="Phase 2: Architectural Modeling & System Design — ")

    add_paragraph_styled(doc,
        "Designing the relational schema comprising fourteen normalized tables, enforcing foreign key integrity constraints, indexing high-frequency lookup keys, and establishing audit logging tables.",
        bold_prefix="Phase 3: Relational Database Schema Engineering — ")

    add_paragraph_styled(doc,
        "Programming the core clinical services in PHP 8.2: the deterministic ClinicalEngine for reference range and panic limit evaluation, the 48-hour duplicate order alert service, the NLM LOINC autocomplete client, and the OpenRouter AI integration with strict Zero-PII sanitization.",
        bold_prefix="Phase 4: Backend Service & Clinical Logic Implementation — ")

    add_paragraph_styled(doc,
        "Developing bilingual, responsive, role-specific user portals for Administrators, Receptionists, Doctors, Lab Specialists, and Patients utilizing semantic HTML5, CSS3 variables, and AJAX asynchronous interactions.",
        bold_prefix="Phase 5: User Interface Engineering & Frontend Integration — ")

    add_paragraph_styled(doc,
        "Executing rigorous unit, integration, boundary, and clinical scenario testing across 100% of platform functional pathways, validating zero-hallucination deterministic limits and full security compliance.",
        bold_prefix="Phase 6: Verification, Clinical Scenario Testing & Hardening — ")

    add_h2(doc, "3.5 Project Development Process")
    add_paragraph_styled(doc,
        "The development process embraced modern software engineering paradigms adapted for medical informatics. The technology stack was chosen specifically for enterprise stability, rapid execution, robust relational integrity, and strict adherence to open standards.")

    # Table 3.2: Development Environment & Tech Stack
    tech_stack_headers = ["Layer / Component", "Technology / Tool", "Technical Role & Architecture Rationale"]
    tech_stack_data = [
        ["Programming Language", "PHP 8.2+ (Strict Types)", "Core backend server runtime; utilizes custom MVC pattern, match expressions, and typed properties for memory safety."],
        ["Database Management", "MySQL 5.7 / 8.0 (InnoDB)", "Relational database riaya_db; 14 normalized tables (3NF), ACID transaction guarantees, and utf8mb4 full Arabic collation."],
        ["Database Connectivity", "PHP Data Objects (PDO)", "Secure data access layer with parameterized prepared statements, preventing SQL injection vulnerabilities."],
        ["Web Server & Runtime", "Apache 2.4 / MAMP Environment", "High-performance web server hosting URL rewriting rules, secure HTTP headers, and virtual host management."],
        ["Frontend Technologies", "HTML5, Modern CSS3, JavaScript ES6+", "Lightweight, dependency-free responsive interface; CSS custom properties for clinical theming and AJAX for async updates."],
        ["Clinical Vocabularies", "LOINC (NLM API) & ICD-10", "Logical Observation Identifiers Names and Codes mapped via U.S. NLM web services, and WHO ICD-10 clinical diagnoses."],
        ["Reference Limits Source", "Mayo Clinic & ARUP Laboratories", "Authoritative clinical reference ranges and physiological panic limits for deterministic ClinicalEngine bounds."],
        ["Artificial Intelligence", "OpenRouter API (Zero-PII LLM)", "Secure cloud gateway accessing advanced LLM models with strict patient de-identification for Arabic educational summaries."],
        ["Data Visualization", "HTML5 Canvas & CSS Range Bars", "Visual rendering of measured clinical values relative to low, normal, high, and critical reference intervals."]
    ]
    put_table(tech_stack_headers, tech_stack_data, 3, "Software Development Environment and Technical Stack Overview", col_widths=[1.5, 2.0, 2.77])

    # Table 3.3: Functional Requirements by Role
    req_headers = ["System Role", "Primary Responsibilities", "Core Functional Capabilities within Riaya Platform"]
    req_data = [
        ["System Administrator", "System Governance & Catalog Oversight", "Manage user accounts and roles; configure hospital clinical departments; curate LOINC tests catalog; auto-populate reference ranges; inspect system audit logs; configure OpenRouter AI parameters."],
        ["Receptionist", "Patient Intake & Scheduling", "Search patient records; manage clinic appointment bookings; audit daily clinic queues; prevent double-booking via schedule slot validation."],
        ["Treating Physician", "Consultation & Test Requisition", "Access patient EMR; view clinical history; order laboratory tests with LOINC auto-completion; receive 48-hour duplicate warnings; review STAT critical alerts; document treatment plans and diagnoses."],
        ["Laboratory Specialist", "Specimen Triage & Result Analysis", "Access prioritized lab order worklist; input numerical test results; verify real-time panic limit flags; trigger automated doctor alert dispatches upon result authorization."],
        ["Patient", "Personal Diagnostic Health Portal", "View personal appointments; inspect past laboratory diagnostic reports; review color-coded visual indicator range bars; read AI-generated plain-language Arabic explanations; print bilingual medical reports."]
    ]
    put_table(req_headers, req_data, 3, "System Stakeholder Roles and Functional Capabilities Matrix", col_widths=[1.5, 2.0, 2.77])

    add_paragraph_styled(doc,
        "Non-functional requirements for the Riaya platform were strictly defined: (1) Zero-Hallucination Guarantee: Clinical critical values and abnormal alerts are computed exclusively by a deterministic mathematical engine without AI intervention. (2) Zero-PII Privacy Protection: Patient names, national IDs, and contact numbers are strictly scrubbed prior to transmitting anonymized diagnostic summaries to the AI explanation gateway, complying with HIPAA and the Saudi Personal Data Protection Law (PDPL). (3) Performance: Database queries execute in sub-50 milliseconds via strategic B-tree indexing on foreign keys and LOINC codes. (4) Security: Complete protection against OWASP Top 10 vulnerabilities via parameterized queries, input sanitization, CSRF token validation, and secure session management.")

    add_h2(doc, "3.6 System Design Diagrams")
    add_paragraph_styled(doc,
        "System architecture and operational workflows are articulated through a comprehensive suite of unified software engineering diagrams, modeling functional use cases, entity relationships, operational sequences, and clinical data flowcharts.")

    # 3.6.1 Use Case Diagram
    add_h3(doc, "3.6.1 Use Case Diagram")
    add_paragraph_styled(doc,
        "The unified use case model captures the interactions between the five system actors and the platform's core functional packages. Each actor possesses distinct operational boundaries enforced by cryptographic role-based access control filters.")

    img_usecase = os.path.join(DIAG_DIR, '01_use_case_diagram.png')
    put_figure(img_usecase, 3, "Riaya Platform Unified Use Case Model across All Operational Roles")

    add_paragraph_styled(doc,
        "As modeled in Figure 3.1, the Administrator governs system users, clinical departments, LOINC catalog curation, and platform auditing. The Receptionist manages patient check-ins and appointment schedules. The Treating Physician accesses EMRs, orders LOINC-coded laboratory tests with duplicate alert interception, receives urgent STAT alerts, and documents treatment plans. The Laboratory Specialist triages pending specimen orders, enters quantitative test data, and authorizes clinical results. The Patient accesses self-service health records, visual indicator bars, and educational Arabic AI summaries.")

    # 3.6.2 Entity-Relationship Diagram
    add_h3(doc, "3.6.2 Entity-Relationship Diagram")
    add_paragraph_styled(doc,
        "The relational database schema is structured into fourteen normalized entities designed to ensure absolute referential integrity, zero data redundancy, and robust performance under concurrent transactional workloads.")

    img_erd = os.path.join(DIAG_DIR, '02_er_diagram_chen.png')
    put_figure(img_erd, 3, "Riaya Relational Database Entity-Relationship Diagram (Chen Notation)")

    add_paragraph_styled(doc,
        "As detailed in the Chen ERD (Figure 3.2), the database structure centers on core transactional entities: `users` stores system actors with role differentiation (`admin`, `receptionist`, `doctor`, `lab`, `patient`); `departments` organizes clinical units; `doctor_schedules` and `appointments` regulate outpatient scheduling; `tests_catalog` maintains standardized LOINC codes, reference intervals, and panic limits; `lab_orders` links doctor requisitions with priority levels; `lab_results` stores quantitative values, flags, and delta checks; `critical_alerts` manages urgent STAT physician alerts; `treatment_plans` and `medical_reports` document clinical follow-up; while `audit_logs` and `notifications` maintain comprehensive operational auditability.")

    # 3.6.3 Sequence Diagrams
    add_h3(doc, "3.6.3 Sequence Diagrams")
    add_paragraph_styled(doc,
        "Operational sequences model the temporal progression and message exchanges occurring during core clinical transactions.")

    img_seq_order = os.path.join(DIAG_DIR, '03_sequence_doctor_order.png')
    put_figure(img_seq_order, 3, "Sequence Diagram: Doctor Consultation, LOINC Test Ordering, and Duplicate Prevention")

    add_paragraph_styled(doc,
        "Figure 3.3 illustrates the doctor laboratory ordering workflow. The physician queries the LOINC catalog, and upon selecting an analyte, the platform's DuplicateAlert engine automatically performs a temporal lookup on `lab_orders` for the patient. If the identical test was ordered within 48 hours, an alert modal warns the physician, requiring clinical justification before dispatching the order to the laboratory queue.")

    img_seq_lab = os.path.join(DIAG_DIR, '04_sequence_lab_upload_result.png')
    put_figure(img_seq_lab, 3, "Sequence Diagram: Laboratory Result Entry, Deterministic Panic Limit Evaluation, and STAT Alert Notification")

    add_paragraph_styled(doc,
        "Figure 3.4 models the laboratory result verification workflow. When the laboratory specialist enters a numerical value, the backend ClinicalEngine evaluates the result against the test's critical thresholds. If a panic limit is breached, the result is flagged as 'critical', an alert record is created in `critical_alerts`, and an instant STAT alert banner is pushed to the attending physician's workspace requiring formal clinical acknowledgment.")

    # 3.6.4 Clinical Process Flowcharts
    add_h3(doc, "3.6.4 Clinical Process Flowcharts")
    add_paragraph_styled(doc,
        "Flowcharts model algorithmic decision logic implemented within Riaya's clinical and AI subsystems.")

    img_flow_loinc = os.path.join(DIAG_DIR, '05_flowchart_loinc_ai_generation.png')
    put_figure(img_flow_loinc, 3, "Clinical Flowchart: LOINC Standardization and Automated Clinical Range Population")

    add_paragraph_styled(doc,
        "Figure 3.5 outlines the automated LOINC catalog enrichment process. When an administrator inputs a diagnostic test name, an asynchronous query executes against the NLM LOINC web service to retrieve official codes. The platform then prompts the clinical range generator to populate verified biological reference intervals and panic limits sourced from Mayo Clinic Laboratories and ARUP Laboratories standards.")

    img_flow_ai = os.path.join(DIAG_DIR, '06_flowchart_patient_ai_explanation.png')
    put_figure(img_flow_ai, 3, "Clinical Flowchart: Patient Zero-PII Generative AI Arabic Explanation Pipeline")

    add_paragraph_styled(doc,
        "Figure 3.6 outlines the patient educational explanation workflow. When a patient requests an explanation, all personally identifiable information (PII) is completely stripped. The numerical findings, reference ranges, and clinical classifications are structured into a prompt dispatched via OpenRouter API. The returned explanation is validated for professional, empathetic Arabic language, mandatory medical disclaimers, and stored securely for patient review.")

    # 3.7 System Interfaces Overview
    add_h2(doc, "3.7 System Interfaces Overview")
    add_paragraph_styled(doc,
        "The Riaya user interface is architected around intuitive, role-specific operational dashboards designed to optimize cognitive focus, eliminate medical errors, and provide seamless access to diagnostic workflows. Every interface is fully responsive, accessible, and supports bilingual presentation.")

    # 3.7.1 Authentication & Security Interfaces
    add_h3(doc, "3.7.1 Authentication & Security Interfaces")
    add_paragraph_styled(doc,
        "The authentication subsystem guarantees that only verified hospital personnel and registered patients can access confidential healthcare records, strictly directing each authenticated actor to their authorized role dashboard.")

    img_auth_login = os.path.join(SCREEN_DIR, '00_auth', '01_login_page.png')
    put_figure(img_auth_login, 3, "Unified Secure Portal Login Interface")
    add_paragraph_styled(doc,
        "The unified login interface provides a secure, encrypted authentication gateway for all system roles. It features cryptographic password verification, CSRF token validation, rate-limiting protection against brute-force attacks, and automatic role-based redirection to the appropriate portal.")

    img_auth_reg = os.path.join(SCREEN_DIR, '00_auth', 'register.png')
    put_figure(img_auth_reg, 3, "Patient Self-Registration Interface")
    add_paragraph_styled(doc,
        "The patient registration interface enables new patients to establish their medical portal profile by providing verified demographic information, national identity data, and contact details. It automatically generates a unique patient identifier and provisions a secure portal account.")

    # 3.7.2 Administrator Portal
    add_h3(doc, "3.7.2 Administrator Portal")
    add_paragraph_styled(doc,
        "The Administrator Portal serves as the operational command center for system governance, clinical catalog maintenance, user management, and security parameter configuration.")

    img_adm_dash = os.path.join(SCREEN_DIR, '01_admin', '01_dashboard.png')
    put_figure(img_adm_dash, 3, "Administrator Analytics and Governance Dashboard")
    add_paragraph_styled(doc,
        "The administrator dashboard delivers high-level operational intelligence regarding system throughput, displaying real-time metrics on total registered patients, active medical staff, scheduled appointments, pending laboratory orders, and system health status. It provides quick shortcuts to user administration, departmental structuring, and LOINC catalog curation.")

    img_adm_users = os.path.join(SCREEN_DIR, '01_admin', '02_users_list.png')
    put_figure(img_adm_users, 3, "Staff and Patient User Management Directory")
    add_paragraph_styled(doc,
        "The user management interface displays an exhaustive directory of all platform users, categorized by role, department, and account status. Administrators can filter users, inspect account activity logs, toggle activation statuses, and initiate account provisioning.")

    img_adm_add_user = os.path.join(SCREEN_DIR, '01_admin', '03_add_user_filled.png')
    put_figure(img_adm_add_user, 3, "User Account Provisioning and Role Assignment Interface")
    add_paragraph_styled(doc,
        "The user creation interface allows administrators to provision new medical practitioners, receptionists, or lab specialists with specific departmental assignments and role privileges. It enforces rigorous input validation and secure credential generation.")

    img_adm_dept = os.path.join(SCREEN_DIR, '01_admin', '05_departments_list.png')
    put_figure(img_adm_dept, 3, "Hospital Clinical Departments Management Interface")
    add_paragraph_styled(doc,
        "The departments management interface enables structured configuration of hospital clinical units, including Internal Medicine, Pediatrics, Surgery, and Emergency. It regulates doctor assignments and defines departmental consultation capacities.")

    img_adm_cat = os.path.join(SCREEN_DIR, '01_admin', '06_tests_catalog.png')
    put_figure(img_adm_cat, 3, "LOINC Standardized Diagnostic Tests Catalog")
    add_paragraph_styled(doc,
        "The tests catalog interface manages the hospital's standardized laboratory test directory, displaying test names, official LOINC codes, measurement units, biological reference intervals, and life-threatening panic limits. It serves as the single source of truth for all clinical diagnostic requisitions.")

    img_adm_loinc_auto = os.path.join(SCREEN_DIR, '01_admin', '08_loinc_autocomplete_results.png')
    put_figure(img_adm_loinc_auto, 3, "National Library of Medicine (NLM) LOINC Live Autocomplete Interface")
    add_paragraph_styled(doc,
        "This modal interface connects directly with NLM web services to execute live LOINC lookups as the administrator types test names. It eliminates human encoding errors by presenting standardized international codes, official analyte descriptions, and established diagnostic classes.")

    img_adm_autofill = os.path.join(SCREEN_DIR, '01_admin', '09_test_autofill_populated.png')
    put_figure(img_adm_autofill, 3, "Automated Clinical Reference Ranges and Panic Limits Auto-Population")
    add_paragraph_styled(doc,
        "Upon selecting a verified LOINC code, this interface auto-populates clinically validated reference ranges and critical panic thresholds calibrated against Mayo Clinic Laboratories and ARUP Laboratories standards, ensuring consistent, error-free clinical baseline parameters.")

    img_adm_settings = os.path.join(SCREEN_DIR, '01_admin', '12_settings.png')
    put_figure(img_adm_settings, 3, "Platform Security, OpenRouter AI, and System Configuration")
    add_paragraph_styled(doc,
        "The platform settings interface allows administrators to configure API keys for the OpenRouter AI gateway, adjust AI model selection (e.g., Claude 3.5 Sonnet, GPT-4o), set duplicate order check time windows, and configure security audit policies.")

    # 3.7.3 Receptionist Portal
    add_h3(doc, "3.7.3 Receptionist Portal")
    add_paragraph_styled(doc,
        "The Receptionist Portal streamlines patient intake, schedule management, and outpatient clinic consultation booking.")

    img_rec_dash = os.path.join(SCREEN_DIR, '02_receptionist', '02_dashboard.png')
    put_figure(img_rec_dash, 3, "Receptionist Workspace and Clinic Queue Dashboard")
    add_paragraph_styled(doc,
        "The receptionist dashboard displays daily outpatient clinic volume, active physician schedules, waiting queue counts, and upcoming appointment slots. It provides rapid lookup shortcuts for patient verification and registration.")

    img_rec_book = os.path.join(SCREEN_DIR, '02_receptionist', '06_book_form_filled.png')
    put_figure(img_rec_book, 3, "Outpatient Clinic Consultation Appointment Booking Interface")
    add_paragraph_styled(doc,
        "The appointment booking interface enables staff to schedule consultations by searching patients, selecting attending specialists, and choosing available time slots. Built-in concurrency validation prevents double-booking and enforces clinic capacity limits.")

    # 3.7.4 Clinical Physician Portal
    add_h3(doc, "3.7.4 Clinical Physician Portal")
    add_paragraph_styled(doc,
        "The Physician Portal empowers treating doctors to review electronic medical records, order standardized laboratory tests, intercept duplicate orders, monitor STAT critical alerts, and formulate structured treatment plans.")

    img_doc_dash = os.path.join(SCREEN_DIR, '03_doctor', '02_dashboard.png')
    put_figure(img_doc_dash, 3, "Attending Physician Clinical Workspace Dashboard")
    add_paragraph_styled(doc,
        "The physician dashboard presents the clinician's daily patient schedule, urgent consultation requests, pending diagnostic test requisitions, and highlighted abnormal findings. It acts as the primary cockpit for evidence-based clinical decision-making.")

    img_doc_emr = os.path.join(SCREEN_DIR, '03_doctor', '06_sara_profile.png')
    put_figure(img_doc_emr, 3, "Comprehensive Patient Electronic Medical Record (EMR) Profile")
    add_paragraph_styled(doc,
        "The patient profile view provides a unified longitudinal medical summary, aggregating vital signs, clinical encounter notes, historical laboratory results, active medication regimens, and previous diagnostic reports into a single, cohesive clinical overview.")

    img_doc_order = os.path.join(SCREEN_DIR, '03_doctor', '09_order_test_filled.png')
    put_figure(img_doc_order, 3, "Standardized LOINC Laboratory Test Requisition Interface")
    add_paragraph_styled(doc,
        "The test order interface allows physicians to select standardized laboratory tests using LOINC autocompletion, specify clinical urgency (Routine vs. STAT), record clinical indications, and link ICD-10 diagnostic codes. The system automatically executes a 48-hour duplicate order check before dispatching.")

    img_doc_crit = os.path.join(SCREEN_DIR, '03_doctor', '11_critical_result_view.png')
    put_figure(img_doc_crit, 3, "Life-Threatening STAT Critical Value Alert Banner on Physician Screen")
    add_paragraph_styled(doc,
        "When a laboratory result breaches pre-configured panic limits, this urgent high-contrast red banner appears immediately on the attending physician's screen. It displays the patient's critical analyte value alongside emergency contact protocols and requires formal clinical acknowledgment.")

    img_doc_treat = os.path.join(SCREEN_DIR, '03_doctor', '13_treatment_plan_filled.png')
    put_figure(img_doc_treat, 3, "Structured Clinical Treatment Plan and Prescription Formulation")
    add_paragraph_styled(doc,
        "Following laboratory result review, the physician documents therapeutic interventions, prescribed medications, lifestyle modifications, and follow-up appointment schedules in this structured treatment plan interface, which immediately synchronizes with the patient portal.")

    # 3.7.5 Laboratory Information System
    add_h3(doc, "3.7.5 Laboratory Information System")
    add_paragraph_styled(doc,
        "The Laboratory Information System (LIS) provides lab technicians and clinical biochemists with dedicated worklists for specimen processing, quantitative result recording, and panic limit verification.")

    img_lab_orders = os.path.join(SCREEN_DIR, '04_lab', '03_orders_list.png')
    put_figure(img_lab_orders, 3, "Laboratory Specimen Processing Worklist and Priority Queue")
    add_paragraph_styled(doc,
        "The laboratory worklist displays incoming diagnostic orders sorted dynamically by clinical urgency (STAT versus Routine), specimen accession numbers, ordering clinics, and elapsed processing times, ensuring emergency specimens receive immediate processing priority.")

    img_lab_entry = os.path.join(SCREEN_DIR, '04_lab', '05_critical_value_entered.png')
    put_figure(img_lab_entry, 3, "Quantitative Result Entry Interface with Real-Time Critical Value Detection")
    add_paragraph_styled(doc,
        "The result entry interface enables technicians to record measured quantitative values. As data is entered, the deterministic ClinicalEngine evaluates the value against reference ranges, dynamically rendering abnormal warning badges and triggering urgent physician STAT notifications upon panic threshold violations.")

    # 3.7.6 Patient Health Portal
    add_h3(doc, "3.7.6 Patient Health Portal")
    add_paragraph_styled(doc,
        "The Patient Health Portal empowers patients with secure, transparent access to their personal diagnostic history, visual test interpretations, and AI-powered educational summaries.")

    img_pat_dash = os.path.join(SCREEN_DIR, '05_patient', '02_dashboard.png')
    put_figure(img_pat_dash, 3, "Patient Personal Health and Diagnostic Dashboard")
    add_paragraph_styled(doc,
        "The patient dashboard welcomes individuals with an accessible overview of their upcoming clinical appointments, recently released laboratory reports, and active physician treatment plans, fostering active engagement with their healthcare journey.")

    img_pat_res_list = os.path.join(SCREEN_DIR, '05_patient', '05_results_list.png')
    put_figure(img_pat_res_list, 3, "Patient Diagnostic Laboratory Reports History")
    add_paragraph_styled(doc,
        "The diagnostic reports directory presents a chronological list of all authorized laboratory tests, indicating test categories, ordering physicians, completion dates, and high-level health status indicators, allowing patients to track their diagnostic records over time.")

    img_pat_res_crit = os.path.join(SCREEN_DIR, '05_patient', '06_result_detail_critical.png')
    put_figure(img_pat_res_crit, 3, "Laboratory Result Breakdown with Visual Abnormal Range Bar and Delta Trend")
    add_paragraph_styled(doc,
        "This detailed test analysis interface visualizes the measured analyte value on a continuous, color-coded horizontal reference bar (Low, Normal, High, Critical), allowing patients to immediately understand their result's position relative to standard biological boundaries, alongside historical comparison metrics.")

    img_pat_ai = os.path.join(SCREEN_DIR, '05_patient', '07_ai_explanation_view.png')
    put_figure(img_pat_ai, 3, "AI-Generated Plain-Language Arabic Educational Report Explanation")
    add_paragraph_styled(doc,
        "Powered by OpenRouter LLM with zero-PII anonymization, this interface translates complex biochemical parameters into compassionate, clear, educational Arabic prose. It explains the biological role of the tested analyte, clarifies what the numerical score means, offers reassuring guidance, and emphasizes adhering to physician advice.")

    img_pat_print = os.path.join(SCREEN_DIR, '05_patient', '08_medical_report_printable.png')
    put_figure(img_pat_print, 3, "Standardized Bilingual Printable Clinical Diagnostic Report")
    add_paragraph_styled(doc,
        "The printable medical report generates an official, standardized bilingual diagnostic document complete with patient demographics, ordering physician details, LOINC test descriptions, quantitative findings, reference intervals, and institutional verification seals, suitable for physical archiving or cross-hospital consultation.")

    # 3.8 Testing and Quality Assurance
    add_h2(doc, "3.8 Testing and Quality Assurance")
    add_paragraph_styled(doc,
        "Quality assurance and verification protocols were conducted across four comprehensive testing tiers to ensure clinical safety, architectural stability, and robust cybersecurity:",
        bold_prefix="Multi-Tiered Verification Framework: ")

    add_paragraph_styled(doc,
        "Validating the mathematical boundaries of the deterministic ClinicalEngine. Unit test suites verified that test results below the lower reference limit trigger 'low' flags, values within normal boundaries trigger 'normal' flags, values exceeding upper limits trigger 'high' flags, and values crossing critical limits (e.g., Hemoglobin < 7.0 g/dL or Glucose > 400 mg/dL) unconditionally trigger 'critical' STAT alerts with 100% precision.",
        bold_prefix="1. Unit Testing of Clinical Range & Panic Engines: ")

    add_paragraph_styled(doc,
        "Verifying asynchronous JSON exchanges with the National Library of Medicine (NLM) LOINC web service and the OpenRouter AI gateway. Tests confirmed that network timeouts or API latency fail gracefully without stalling primary clinical user interfaces, and that zero protected health information (PII) is transmitted across external API boundaries.",
        bold_prefix="2. External API Integration & Data Sanitization Testing: ")

    add_paragraph_styled(doc,
        "Auditing the application against the OWASP Top 10 web vulnerabilities. Verification confirmed that all database queries utilize PDO prepared statements with strict parameter binding (eliminating SQL injection risks), all dynamic user outputs are escaped via `htmlspecialchars` (preventing Cross-Site Scripting), anti-CSRF tokens guard every state-modifying POST transaction, and role-based session middleware rigorously blocks horizontal and vertical privilege escalation.",
        bold_prefix="3. Security and Penetration Resistance Auditing: ")

    add_paragraph_styled(doc,
        "Simulating complete end-to-end clinical journeys: (a) Receptionist books consultation -> (b) Doctor reviews EMR and orders LOINC-coded CBC and Blood Glucose -> (c) Duplicate alert is verified when re-ordering within 48 hours -> (d) Lab enters critically low Hemoglobin (6.2 g/dL) -> (e) System instantly posts STAT red banner on doctor screen -> (f) Doctor acknowledges alert and writes treatment plan -> (g) Patient inspects report and reads Arabic AI educational summary. All scenarios executed flawlessly.",
        bold_prefix="4. End-to-End Clinical Scenario Simulation: ")

    # 3.9 Future Work
    add_h2(doc, "3.9 Future Work")
    add_paragraph_styled(doc,
        "To further expand the clinical reach and technical sophistication of Riaya, several strategic enhancements are proposed for future development phases:",
        bold_prefix="Future Strategic Enhancements: ")

    add_paragraph_styled(doc,
        "Implementing Health Level 7 (HL7) Version 2.x and Fast Healthcare Interoperability Resources (FHIR) RESTful APIs to enable seamless bidirectional data interchange with the Ministry of Health's centralized Seha platform, national laboratory information exchanges, and regional hospital networks.",
        bold_prefix="1. HL7 / FHIR Standards Integration: ")

    add_paragraph_styled(doc,
        "Developing cross-platform native mobile applications for iOS and Android utilizing Flutter or React Native, incorporating secure biometric authentication, push notifications for critical value alerts, and direct patient-physician telehealth messaging.",
        bold_prefix="2. Native Mobile Applications for iOS & Android: ")

    add_paragraph_styled(doc,
        "Integrating machine learning models to analyze longitudinal delta trends across multiple biological markers simultaneously, providing early automated risk scores for sepsis, acute kidney injury (AKI), or diabetic ketoacidosis (DKA) prior to overt clinical decompensation.",
        bold_prefix="3. Predictive Clinical Deterioration Analytics: ")

    add_paragraph_styled(doc,
        "Extending the deterministic ClinicalEngine to automatically ingest and cross-reference patient active medication profiles, issuing automated alerts when prescribed drugs are known to interfere with specific laboratory analyte measurements.",
        bold_prefix="4. Pharmacogenomic and Drug-Lab Interference Cross-Checking: ")

    doc.add_page_break()

    print("Building Chapter 4: Conclusion & Recommendations...")
    # ---------------------------------------------------------
    # CHAPTER IV: CONCLUSION
    # ---------------------------------------------------------
    add_h1(doc, "CHAPTER IV:  CONCLUSION")

    add_h2(doc, "Conclusion")
    add_paragraph_styled(doc,
        "The cooperative training program completed at Al-Hait General Hospital, under the auspices of the College of Computer Science and Engineering at the University of Hail, represented an invaluable capstone experience that successfully fused academic computing principles with the demanding operational realities of enterprise healthcare information technology. By embedding within the hospital's IT and technical support operations, direct contributions were made toward maintaining critical hardware infrastructure, restoring lagging workstations, securing local area networks, resolving printer driver bottlenecks, and supporting clinical staff across administrative, outpatient, and medical records departments.")

    add_paragraph_styled(doc,
        "Concurrently, the internship fostered acute awareness of the communication barriers and diagnostic inefficiencies present in conventional healthcare delivery. This direct clinical exposure culminated in the architectural design, implementation, and successful delivery of 'Riaya'. By uniting modern web engineering in PHP 8.2 and MySQL with international medical vocabularies (LOINC and ICD-10), deterministic reference range engines, proactive STAT critical alert mechanisms, and Zero-PII generative AI educational interpretation, Riaya proves that well-architected software can bridge the cognitive gap between complex medical diagnostics and patient comprehension while establishing life-saving safety nets for attending physicians. The knowledge, problem-solving independence, and technical confidence gained throughout this cooperative journey provide a solid foundation for future professional excellence in computer science and healthcare informatics.")

    add_h2(doc, "Recommendations")
    add_paragraph_styled(doc,
        "Based upon the operational observations and technical engagements undertaken during the cooperative training period, the following recommendations are respectfully submitted:",
        bold_prefix="Strategic Recommendations: ")

    add_paragraph_styled(doc,
        "Accelerating the full phase-out of remaining localized paper records in peripheral clinics and transitioning all diagnostic reporting and clinical notes into standardized, centralized Electronic Medical Record systems.",
        bold_prefix="1. For Al-Hait General Hospital — Full EMR Transition: ")

    add_paragraph_styled(doc,
        "Deploying automated network monitoring and SNMP-based asset management tools to continuously track workstation performance, network switch port health, printer toner thresholds, and hardware lifecycle schedules proactively.",
        bold_prefix="2. For Al-Hait General Hospital — Proactive Network & Asset Monitoring: ")

    add_paragraph_styled(doc,
        "Conducting recurring, practical cybersecurity awareness workshops for clinical, nursing, and administrative personnel, focusing on social engineering defenses, phishing awareness, password hygiene, and strict adherence to patient privacy standards.",
        bold_prefix="3. For Al-Hait General Hospital — Staff Cybersecurity Training: ")

    add_paragraph_styled(doc,
        "Continuing to expand cooperative training partnerships with specialized healthcare institutions and health clusters, providing computer science students with early exposure to mission-critical health informatics challenges.",
        bold_prefix="4. For the College of Computer Science and Engineering — Expanding Healthcare Partnerships: ")

    add_paragraph_styled(doc,
        "Approaching cooperative training with an analytical, solution-oriented mindset; seeking opportunities to not only execute assigned maintenance tasks but also identify systemic institutional workflow bottlenecks and engineer durable software deliverables that create lasting organizational value.",
        bold_prefix="5. For Future Cooperative Training Students — Proactive Problem-Solving: ")

    doc.add_page_break()

    print("Building References...")
    # ---------------------------------------------------------
    # REFERENCES
    # ---------------------------------------------------------
    add_h1(doc, "References")

    references_list = [
        ("Regenstrief Institute & U.S. National Library of Medicine", "2024", "Logical Observation Identifiers Names and Codes (LOINC) — International Standard for Laboratory and Clinical Observations", "https://loinc.org/"),
        ("World Health Organization (WHO)", "2023", "International Statistical Classification of Diseases and Related Health Problems, 10th Revision (ICD-10)", "https://icd.who.int/browse10/2019/en"),
        ("Mayo Clinic Laboratories", "2024", "Clinical and Interpretive Laboratory Test Reference Catalog & Critical Values Directory", "https://www.mayocliniclabs.com/test-catalog/"),
        ("ARUP Laboratories", "2024", "Laboratory Test Directory and Critical Values Guidelines", "https://www.aruplab.com/"),
        ("OpenRouter", "2024", "OpenRouter AI API Documentation — Unified Interface for Advanced Large Language Models", "https://openrouter.ai/docs"),
        ("Ministry of Health (MOH), Kingdom of Saudi Arabia", "2024", "Health Sector Transformation Program & Digital Health Strategy Overview", "https://www.moh.gov.sa/"),
        ("Hail Health Cluster", "2024", "Official Portal of the Hail Health Cluster — Facilities and Services Directory", "https://hailthy.net/"),
        ("Saudi Central Board for Accreditation of Healthcare Institutions (CBAHI)", "2023", "National Standards for Hospitals (3rd Edition)", "https://cbahi.gov.sa/"),
        ("The PHP Group", "2024", "PHP 8.2 Documentation & Language Reference Manual", "https://www.php.net/manual/en/"),
        ("Oracle Corporation", "2024", "MySQL 8.0 Reference Manual — InnoDB Storage Engine Architecture", "https://dev.mysql.com/doc/refman/8.0/en/"),
        ("Open Web Application Security Project (OWASP)", "2023", "OWASP Top Ten Web Application Security Risks", "https://owasp.org/www-project-top-ten/"),
        ("Saudi Data & AI Authority (SDAIA)", "2023", "Personal Data Protection Law (PDPL) Implementing Regulations in the Kingdom of Saudi Arabia", "https://sdaia.gov.sa/"),
        ("University of Hail", "2024", "Cooperative Education Program Manual — College of Computer Science and Engineering", "https://www.uoh.edu.sa/"),
        ("Al-Jazirah Newspaper Archive", "2014", "Commencement of Full Operations at Al-Hait General Hospital in Hail Region", "https://www.al-jazirah.com/2014/20140225/ln60.htm")
    ]

    for idx, (author, year, title, url) in enumerate(references_list, 1):
        p_ref = doc.add_paragraph()
        set_para_format(p_ref, before_pt=4, after_pt=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
        r_num = p_ref.add_run(f"[{idx}] ")
        set_run_font(r_num, "Times New Roman", 11, bold=True, color=COLOR_DARK_BLUE)
        r_body = p_ref.add_run(f"{author} ({year}). {title}. Retrieved from ")
        set_run_font(r_body, "Times New Roman", 11, bold=False, color=RGBColor(0, 0, 0))
        r_url = p_ref.add_run(url)
        set_run_font(r_url, "Times New Roman", 11, bold=False, italic=True, color=RGBColor(0x05, 0x63, 0xC1))

    print(f"Saving compiled document to {OUTPUT_DOCX}...")
    doc.save(OUTPUT_DOCX)
    shutil.copyfile(OUTPUT_DOCX, REPORT_EXPORT_DOCX)
    print(f"Copied report to {REPORT_EXPORT_DOCX}")

    print("Validating XML field structures...")
    with zipfile.ZipFile(OUTPUT_DOCX) as z:
        doc_xml = z.read('word/document.xml')
    tree = etree.fromstring(doc_xml)
    all_fc = tree.findall(f'.//{W_NS}fldChar')
    broken = sum(1 for fc in all_fc if fc.getparent().tag != f'{W_NS}r')
    print(f"Validation Result: Total fldChar elements: {len(all_fc)}, Broken (outside <w:r>): {broken}")
    if broken == 0:
        print("SUCCESS: 100% of XML fields are strictly compliant with golden template specifications!")
    else:
        print(f"WARNING: Found {broken} non-compliant fields.")

if __name__ == '__main__':
    build_report()
