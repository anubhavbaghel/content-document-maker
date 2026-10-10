#!/usr/bin/env python3
"""
generate_doc.py - Generates a standard Word document
directly from content.json, with full typography, styled header tables, USP tables,
dynamic service blocks, FAQs, and CTA callouts.
"""

import html
import itertools
import json
import re
from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

INTEXT_BUTTON_POOL = [
    "Contact Us",
    "Contact Us Now",
    "Contact Us Today",
    "Enquire Now",
    "Contact Our Team",
    "Contact Our Experts",
    "Enquire Today",
]

LEAD_BUTTON_POOL = [
    "Get In Touch",
    "Enquire Now",
    "Get A Quote",
    "Have Any Questions?",
    "Find Out More",
    "Enquire Today",
    "Reach Your Local Expert",
    "Start Your Enquiry",
    "Ask A Question",
]


def make_picker(pool):
    cycle = itertools.cycle(pool)

    def pick(avoid=None):
        val = next(cycle)
        if avoid is not None and val == avoid:
            val = next(cycle)
        return val

    return pick


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


BOLD_MARKUP_RE = re.compile(r'\*\*(.+?)\*\*')


def add_formatted_paragraph(doc_or_cell, text, bold_default=False, font_size=11, italic=False, color=None, space_after=6):
    if hasattr(doc_or_cell, 'add_paragraph'):
        p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()
    
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15

    if not text:
        return p

    lines = text.split("\n")
    for l_idx, line in enumerate(lines):
        if l_idx > 0:
            if hasattr(doc_or_cell, 'add_paragraph'):
                p = doc_or_cell.add_paragraph()
            else:
                p = doc_or_cell.add_paragraph()
            p.paragraph_format.space_after = Pt(space_after)
            p.paragraph_format.line_spacing = 1.15

        pos = 0
        matches = list(BOLD_MARKUP_RE.finditer(line))
        if not matches:
            r = p.add_run(line)
            r.font.name = "Calibri"
            r.font.size = Pt(font_size)
            r.font.bold = bold_default
            r.font.italic = italic
            if color:
                r.font.color.rgb = color
        else:
            for m in matches:
                if m.start() > pos:
                    r = p.add_run(line[pos:m.start()])
                    r.font.name = "Calibri"
                    r.font.size = Pt(font_size)
                    r.font.bold = bold_default
                    r.font.italic = italic
                    if color:
                        r.font.color.rgb = color
                r_bold = p.add_run(m.group(1))
                r_bold.font.name = "Calibri"
                r_bold.font.size = Pt(font_size)
                r_bold.font.bold = True
                r_bold.font.italic = italic
                if color:
                    r_bold.font.color.rgb = color
                pos = m.end()
            if pos < len(line):
                r = p.add_run(line[pos:])
                r.font.name = "Calibri"
                r.font.size = Pt(font_size)
                r.font.bold = bold_default
                r.font.italic = italic
                if color:
                    r.font.color.rgb = color
    return p


def add_heading_p(doc, text, level=2, space_before=10, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.bold = True
    if level == 1:
        r.font.size = Pt(15)
        r.font.color.rgb = RGBColor(31, 78, 121)
    elif level == 2:
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(46, 117, 182)
    elif level == 3:
        r.font.size = Pt(11.5)
        r.font.color.rgb = RGBColor(68, 84, 106)
    return p


def build_front_matter(doc, data, page_names):
    # Front-matter info table
    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "A6B9D0", sz="6")
    
    # Header row
    hdr = table.rows[0]
    set_cell_background(hdr.cells[0], "1F4E79")
    set_cell_background(hdr.cells[1], "1F4E79")
    add_formatted_paragraph(hdr.cells[0], "COMPANY / CLIENT DETAILS", bold_default=True, font_size=12, color=RGBColor(255, 255, 255))
    add_formatted_paragraph(hdr.cells[1], "OPENING HOURS & QA FEEDBACK", bold_default=True, font_size=12, color=RGBColor(255, 255, 255))
    
    # Row 1: Company details & Opening hours
    row1 = table.rows[1]
    set_cell_margins(row1.cells[0], 120, 120, 150, 150)
    set_cell_margins(row1.cells[1], 120, 120, 150, 150)
    
    contact_p = data["pages"][1] if len(data["pages"]) > 1 else data["pages"][0]
    instructions = contact_p.get("PageData", {}).get("instructions", "")
    
    co_details = (
        "Company Name: The Grange Retirement Home\n\n"
        "Address:\nRuxbury Road, St. Ann's Hill\nChertsey, Surrey, KT16 9EP\n\n"
        "Email: info@thegrangechertsey.co.uk\n"
        "Brochure: https://www.thegrangechertsey.co.uk/brochure/\n\n"
        "Social Media & Links:\n"
        "• LinkedIn | CQC | Google My Business | Facebook\n"
        "• Carehome.co.uk Profile: Listed"
    )
    add_formatted_paragraph(row1.cells[0], co_details, font_size=10.5)

    hours_details = (
        "Visiting Hours:\n"
        "Open door visitor policy — Visits warmly welcomed\n\n"
        "Care Coverage:\n"
        "24 Hours a day / 7 Days a week (Qualified nurses on duty)\n\n"
        "QA Feedback:\n"
        "Checked against CQC requirements and person-centred care guidelines."
    )
    add_formatted_paragraph(row1.cells[1], hours_details, font_size=10.5)

    # Row 2: Sitemap & Notes
    row2 = table.rows[2]
    set_cell_margins(row2.cells[0], 120, 120, 150, 150)
    set_cell_margins(row2.cells[1], 120, 120, 150, 150)
    
    sitemap_text = "SITE MAP:\n\n" + "\n".join(page_names)
    add_formatted_paragraph(row2.cells[0], sitemap_text, font_size=10.5)
    
    notes_text = (
        "Notes for the QA:\n"
        "• Good CQC rating, Low staff turnover, Home cooked meals, and 24h nursing featured throughout.\n\n"
        "Notes for the Designer:\n"
        "• Ensure easy access to enquiry forms and brochure downloads.\n"
        "• Feature images of landscaped gardens and en suite rooms."
    )
    add_formatted_paragraph(row2.cells[1], notes_text, font_size=10.5)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(18)


def render_page_header_table(doc, page_label, meta_heading, meta_description, headline):
    table = doc.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "B4C6E7", sz="4")

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Inches(1.8)
        row.cells[1].width = Inches(4.7)
        set_cell_margins(row.cells[0], 80, 80, 120, 120)
        set_cell_margins(row.cells[1], 80, 80, 120, 120)

    # Row 0: Page Label
    set_cell_background(table.rows[0].cells[0], "2F5597")
    set_cell_background(table.rows[0].cells[1], "2F5597")
    add_formatted_paragraph(table.rows[0].cells[0], page_label, bold_default=True, font_size=12, color=RGBColor(255, 255, 255))
    add_formatted_paragraph(table.rows[0].cells[1], "", font_size=10)

    # Row 1: Page/Meta Title
    set_cell_background(table.rows[1].cells[0], "D9E1F2")
    add_formatted_paragraph(table.rows[1].cells[0], "Page/Meta Title", bold_default=True, font_size=10.5)
    add_formatted_paragraph(table.rows[1].cells[1], meta_heading or "", font_size=10.5)

    # Row 2: Meta Description
    set_cell_background(table.rows[2].cells[0], "F2F2F2")
    add_formatted_paragraph(table.rows[2].cells[0], "Meta Description", bold_default=True, font_size=10.5)
    add_formatted_paragraph(table.rows[2].cells[1], meta_description or "", font_size=10.5)

    # Row 3: H1 (Hero image text)
    set_cell_background(table.rows[3].cells[0], "D9E1F2")
    add_formatted_paragraph(table.rows[3].cells[0], "H1 (Hero image text)", bold_default=True, font_size=10.5)
    add_formatted_paragraph(table.rows[3].cells[1], headline or "", bold_default=True, font_size=11)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def render_usps_table(doc, usps):
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "D9D9D9", sz="4")
    
    # Header
    set_cell_background(table.rows[0].cells[0], "41719C")
    set_cell_background(table.rows[0].cells[1], "41719C")
    set_cell_background(table.rows[0].cells[2], "41719C")
    
    add_formatted_paragraph(table.rows[0].cells[0], "USP / Key Highlight 1", bold_default=True, font_size=10, color=RGBColor(255, 255, 255))
    add_formatted_paragraph(table.rows[0].cells[1], "USP / Key Highlight 2", bold_default=True, font_size=10, color=RGBColor(255, 255, 255))
    add_formatted_paragraph(table.rows[0].cells[2], "USP / Key Highlight 3", bold_default=True, font_size=10, color=RGBColor(255, 255, 255))

    for idx, (cell, val) in enumerate(zip(table.rows[1].cells, usps)):
        cell.width = Inches(2.1)
        set_cell_margins(cell, 80, 80, 100, 100)
        set_cell_background(cell, "F8F9FA" if idx % 2 == 0 else "FFFFFF")
        add_formatted_paragraph(cell, f"• {val}" if val else "", bold_default=True, font_size=10.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)


def render_cta_table(doc, heading, body_text, button_text):
    table = doc.add_table(rows=2, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "2F5597", sz="8")
    
    set_cell_background(table.rows[0].cells[0], "2F5597")
    add_formatted_paragraph(table.rows[0].cells[0], "CALL TO ACTION (CTA BANNER)", bold_default=True, font_size=11, color=RGBColor(255, 255, 255))
    
    content_cell = table.rows[1].cells[0]
    set_cell_margins(content_cell, 120, 120, 150, 150)
    set_cell_background(content_cell, "F2F4F8")
    
    cta_content = (
        f"**Title:** {heading or 'Get In Touch'}\n"
        f"**Text:** {body_text or 'Call us today to discuss care options.'}\n"
        f"**Button:** [{button_text or 'Contact Us Today'}] > Link to Contact Form"
    )
    add_formatted_paragraph(content_cell, cta_content, font_size=10.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(14)


def build_page_home(doc, page, page_label, pick_lead, pick_intext):
    content = page.get("AIContent", {})
    
    render_page_header_table(
        doc, page_label,
        content.get("meta_data", {}).get("heading"),
        content.get("meta_data", {}).get("description"),
        content.get("headline")
    )
    
    # Subheadline / SQOT
    if content.get("subheadline"):
        add_formatted_paragraph(doc, f"**Key Strengths:** {content.get('subheadline')}", italic=True, font_size=11, color=RGBColor(80, 80, 80))
    
    # Hero Button
    hero_btn = content.get("hero_cta") or pick_lead()
    add_formatted_paragraph(doc, f"Hero image button: [{hero_btn}] > Link to Enquiries", bold_default=True, font_size=10.5, color=RGBColor(46, 117, 182))
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # USPs
    render_usps_table(doc, [content.get("USP1"), content.get("USP2"), content.get("USP3")])

    # Section 1
    sec1 = content.get("section1", {})
    if sec1:
        add_heading_p(doc, sec1.get("heading", ""), level=2)
        add_formatted_paragraph(doc, sec1.get("description", ""))

    # Services section
    svc_title = content.get("service_title") or "The Care and Support We Provide"
    add_heading_p(doc, svc_title, level=2)
    
    for i in range(1, 7):
        svc = content.get(f"service_0{i}")
        if isinstance(svc, dict):
            add_heading_p(doc, svc.get("headline", ""), level=3)
            add_formatted_paragraph(doc, svc.get("summary", ""))
            add_formatted_paragraph(doc, f"[{svc.get('headline', '')}] > Relevant Service Page", italic=True, color=RGBColor(46, 117, 182))

    # Midpage CTA
    add_formatted_paragraph(doc, f"[{pick_intext()}] > Button", bold_default=True, color=RGBColor(46, 117, 182))

    # Section 2: Why choose us
    sec2 = content.get("section2", {})
    if sec2:
        add_heading_p(doc, sec2.get("heading", "Why Choose Us"), level=2)
        add_formatted_paragraph(doc, sec2.get("description", ""))

    # Section 3: Adapting support
    sec3 = content.get("section3", {})
    if sec3:
        add_heading_p(doc, sec3.get("heading", ""), level=2)
        add_formatted_paragraph(doc, sec3.get("description", ""))

    # Process Steps
    proc_title = content.get("process_title") or "How to take the next step with us"
    add_heading_p(doc, proc_title, level=2)
    for key in ("process_1", "process_2", "process_3"):
        p_item = content.get(key)
        if isinstance(p_item, dict):
            add_heading_p(doc, p_item.get("heading", ""), level=3)
            add_formatted_paragraph(doc, p_item.get("description", ""))

    # Propositions
    prop_title = content.get("proposition_title") or "What makes life here feel more settled"
    add_heading_p(doc, prop_title, level=2)
    for key in ("proposition_1", "proposition_2", "proposition_3"):
        p_item = content.get(key)
        if isinstance(p_item, dict):
            add_heading_p(doc, p_item.get("heading", ""), level=3)
            add_formatted_paragraph(doc, p_item.get("description", ""))

    # Midpage CTA 2
    add_formatted_paragraph(doc, f"[{pick_intext()}] > Button", bold_default=True, color=RGBColor(46, 117, 182))

    # CTA Table
    cta = content.get("CTA", {})
    render_cta_table(doc, cta.get("heading"), cta.get("action"), pick_lead(avoid=hero_btn))
    doc.add_page_break()


def build_page_generic(doc, page, page_label, pick_lead, pick_intext):
    content = page.get("AIContent", {})
    
    render_page_header_table(
        doc, page_label,
        content.get("meta_data", {}).get("heading"),
        content.get("meta_data", {}).get("description"),
        content.get("headline")
    )

    if content.get("subheadline"):
        add_formatted_paragraph(doc, f"**Highlights:** {content.get('subheadline')}", italic=True, font_size=11, color=RGBColor(80, 80, 80))

    hero_btn = content.get("hero_cta") or pick_lead()
    add_formatted_paragraph(doc, f"Hero image button: [{hero_btn}] > Link to Enquiries", bold_default=True, font_size=10.5, color=RGBColor(46, 117, 182))
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    render_usps_table(doc, [content.get("USP1"), content.get("USP2"), content.get("USP3")])

    # Sections 1 to 4
    for i in range(1, 5):
        sec = content.get(f"section{i}")
        if isinstance(sec, dict):
            add_heading_p(doc, sec.get("heading", ""), level=2)
            add_formatted_paragraph(doc, sec.get("description", ""))

    # In-text CTA Button
    add_formatted_paragraph(doc, f"[{pick_intext()}] > Button", bold_default=True, color=RGBColor(46, 117, 182))

    # Propositions
    prop_title = content.get("proposition_title")
    if prop_title:
        add_heading_p(doc, prop_title, level=2)
    for key in ("proposition1", "proposition2", "proposition3", "proposition_1", "proposition_2", "proposition_3"):
        prop = content.get(key)
        if isinstance(prop, dict):
            add_heading_p(doc, prop.get("heading", ""), level=3)
            add_formatted_paragraph(doc, prop.get("description", ""))

    # FAQs
    faqs = [content.get(f"FAQs{i}") for i in (1, 2, 3)]
    faqs = [f for f in faqs if isinstance(f, dict)]
    if faqs:
        add_heading_p(doc, "Frequently Asked Questions (FAQs)", level=2)
        for faq in faqs:
            add_heading_p(doc, f"Q: {faq.get('question', '')}", level=3)
            add_formatted_paragraph(doc, faq.get("answer", ""))

    # Bottom CTA Table
    cta = content.get("CTA", {})
    render_cta_table(doc, cta.get("heading"), cta.get("action"), pick_lead(avoid=hero_btn))
    doc.add_page_break()


def build_page_contact(doc, page, page_label, pick_lead):
    content = page.get("AIContent", {})
    
    render_page_header_table(
        doc, page_label,
        content.get("meta_data", {}).get("heading"),
        content.get("meta_data", {}).get("description"),
        content.get("headline")
    )

    if content.get("website_conclusion"):
        add_formatted_paragraph(doc, content.get("website_conclusion"))

    hero_btn = content.get("hero_cta") or pick_lead()
    add_formatted_paragraph(doc, f"Hero image button: [{hero_btn}] > Link to form/email", bold_default=True, font_size=10.5, color=RGBColor(46, 117, 182))
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    sec1 = content.get("section1", {})
    if isinstance(sec1, dict):
        add_heading_p(doc, sec1.get("heading", "Get In Touch"), level=2)
        add_formatted_paragraph(doc, sec1.get("description", ""))

    usps = [content.get("USP1"), content.get("USP2"), content.get("USP3")]
    if any(usps):
        render_usps_table(doc, usps)

    # Contact form blueprint
    add_heading_p(doc, "Contact Form Blueprint", level=2)
    form_text = (
        "Form Fields:\n"
        "• Full Name *\n"
        "• Email Address *\n"
        "• Phone Number *\n"
        "• Care Type Needed (Dropdown: Residential Care, 24 Hour Nursing, Respite / Day Care, General Enquiry) *\n"
        "• Message / Specific Requirements\n"
        "• [Submit Enquiry] > info@thegrangechertsey.co.uk"
    )
    add_formatted_paragraph(doc, form_text, font_size=10.5)


def generate_document(json_path, output_path):
    with open(json_path, encoding='utf-8') as f:
        data = json.load(f)
    
    pages = data["pages"]
    home_pages = [p for p in pages if p.get("PageType") == "home"]
    contact_pages = [p for p in pages if p.get("PageType") == "contact"]
    other_pages = [p for p in pages if p.get("PageType") not in ("home", "contact")]

    pick_lead = make_picker(LEAD_BUTTON_POOL)
    pick_intext = make_picker(INTEXT_BUTTON_POOL)

    doc = docx.Document()
    
    # Page Setup (Standard A4 / 1-inch margins)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    sitemap_names = (
        (["Home (Page 1)"] if home_pages else []) +
        [f"{p.get('PageName')} (Page {i+2})" for i, p in enumerate(other_pages)] +
        [f"Contact Us (Page {len(other_pages)+2})"] if contact_pages else []
    )

    build_front_matter(doc, data, sitemap_names)

    page_num = 1
    if home_pages:
        build_page_home(doc, home_pages[0], f"Home (Page {page_num})", pick_lead, pick_intext)
        page_num += 1

    for p in other_pages:
        label = f"{p.get('PageName')} (Page {page_num})"
        build_page_generic(doc, p, label, pick_lead, pick_intext)
        page_num += 1

    if contact_pages:
        label = f"Contact Us (Page {page_num})"
        build_page_contact(doc, contact_pages[0], label, pick_lead)

    doc.save(output_path)
    print(f"Successfully generated: {output_path}")


if __name__ == "__main__":
    generate_document("content.json", "The_Grange_Retirement_Home_Document.docx")
