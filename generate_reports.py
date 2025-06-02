from docx import Document as DocxDocument
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import os
import re
import argparse

# Map country to flag image filenames (ensure these files exist)
country_flags = {
    'Luxembourg': 'flags/luxembourg.png',
    'Ireland': 'flags/ireland.png',
    'UK': 'flags/uk.png',
    'Switzerland': 'flags/switzerland.png',
    'European Union': 'flags/european_union.png'
}

styles = getSampleStyleSheet()
normal_style = styles['Normal']
normal_style.fontName = 'Helvetica'
normal_style.fontSize = 10
normal_style.leading = 12

label_style = ParagraphStyle(
    'LabelStyle',
    parent=normal_style,
    backColor=colors.lightblue,
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=1,  # center align
    spaceAfter=4,
)

value_style = ParagraphStyle(
    'ValueStyle',
    parent=normal_style,
    fontName='Helvetica',
    fontSize=10,
    leading=12,
)

def parse_docx(file_path):
    print(f"[parse_docx] Loading document: {file_path}")
    document = DocxDocument(file_path)
    full_text = ' '.join([p.text.strip() for p in document.paragraphs if p.text.strip()])
    print(f"[parse_docx] Full text length: {len(full_text)} chars")

    raw_entries = re.split(r'(?=Title:)', full_text)
    print(f"[parse_docx] Found {len(raw_entries)} raw entries (including empty first)")

    entries = []
    for idx, entry_text in enumerate(raw_entries):
        entry_text = entry_text.strip()
        if not entry_text:
            continue
        print(f"[parse_docx] Parsing entry #{idx+1}")

        pattern = (
            r'Title:\s*(.*?)\s+Date:\s*(.*?)\s+Country:\s*(.*?)\s+Summary:\s*(.*?)\s+Link:\s*(.*?)\s+Availability:\s*(.*)'
        )
        match = re.match(pattern, entry_text, re.DOTALL)
        if not match:
            print(f"  [WARNING] Entry #{idx+1} does not match expected format!")
            print(f"  Text: {entry_text[:100]}...")
            continue
        
        entry = {
            "Title": match.group(1).strip(),
            "Date": match.group(2).strip(),
            "Country": match.group(3).strip(),
            "Summary": match.group(4).strip(),
            "Link": match.group(5).strip(),
            "Availability": match.group(6).strip(),
        }

        print(f"  Parsed fields:")
        for k,v in entry.items():
            print(f"    {k}: {v[:50]}{'...' if len(v)>50 else ''}")
        entries.append(entry)

    print(f"[parse_docx] Completed parsing entries. Total: {len(entries)}")
    return entries

def build_table_for_entry(entry):
    print(f"[build_table_for_entry] Building table for entry: {entry['Title']}")

    # Prepare Title split for two lines max
    title = entry["Title"]
    if len(title) > 40:
        parts = title.split(' ')
        mid = len(parts)//2
        title_text = ' '.join(parts[:mid]) + '\n' + ' '.join(parts[mid:])
    else:
        title_text = title

    # Prepare flag image if exists
    flag_path = country_flags.get(entry["Country"])
    flag_img = None
    if flag_path and os.path.isfile(flag_path):
        try:
            if entry["Country"] == "Switzerland":
                # Square flag for Switzerland
                flag_img = Image(flag_path, width=0.5*inch, height=0.5*inch)
            else:
                # Wider flags for others, keep aspect ratio approx 5:3
                flag_img = Image(flag_path, width=0.5*inch, height=0.3*inch)
        except Exception as e:
            print(f"  Error loading flag image: {e}")
            flag_img = None
    else:
        print(f"  No flag image found for country '{entry['Country']}' at '{flag_path}'")

    # Wrap flag image in a Table cell for centering
    if flag_img:
        flag_img = Table([[flag_img]], colWidths=[0.5*inch], rowHeights=[0.5*inch])
        flag_img.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))

    # Increase font leading for multiline wrapping text to avoid overlap
    summary_style = ParagraphStyle(
        'SummaryStyle',
        parent=value_style,
        leading=16,
    )

    # Table data (6 columns)
    data = [
        [
            Paragraph('<b>Title</b>', label_style),
            Paragraph(title_text, value_style),
            Paragraph('<b>Date</b>', label_style),
            Paragraph(entry['Date'], value_style),
            Paragraph('<b>Country</b>', label_style),
            flag_img if flag_img else Paragraph(entry['Country'], value_style)
        ],
        [Paragraph('<b>Summary</b>', label_style), Paragraph(entry['Summary'], summary_style), '', '', '', ''],
        [Paragraph('<b>Link</b>', label_style), Paragraph(entry['Link'], value_style), '', '', '', ''],
        [Paragraph('<b>Availability</b>', label_style), Paragraph(entry['Availability'], value_style), '', '', '', ''],
    ]

    # Set row heights: fixed except summary auto
    base_height = 14
    row_heights = [
        base_height * 3,  # Title row
        None,             # Summary auto height
        base_height * 3,  # Link row
        base_height * 3   # Availability row
    ]

    t = Table(data, colWidths=[1*inch, 2.5*inch, 0.8*inch, 1*inch, 0.8*inch, 1*inch], rowHeights=row_heights)
    style = TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.lightblue),
        ('BACKGROUND', (2,0), (2,0), colors.lightblue),
        ('BACKGROUND', (4,0), (4,0), colors.lightblue),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,0), (1,0), 'LEFT'),
        ('ALIGN', (1,1), (1,3), 'LEFT'),
        ('ALIGN', (5,0), (5,0), 'CENTER'),
        ('SPAN', (1,1), (-1,1)),
        ('SPAN', (1,2), (-1,2)),
        ('SPAN', (1,3), (-1,3)),
        ('BACKGROUND', (0,1), (0,1), colors.lightblue),
        ('BACKGROUND', (0,2), (0,2), colors.lightblue),
        ('BACKGROUND', (0,3), (0,3), colors.lightblue),
    ])
    t.setStyle(style)

    print(f"  Table built for entry: {entry['Title']}")
    return t


def set_cell_background(cell, color_hex):
    if color_hex is None:
        return
    shading_elm = parse_xml(
        r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex)
    )
    tc_pr = cell._tc.get_or_add_tcPr()
    # Remove existing shading if any
    for child in tc_pr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd'):
        tc_pr.remove(child)
    tc_pr.append(shading_elm)


def create_word(entries, output_docx):
    print("[create_word] Creating Word document:", output_docx)
    doc = DocxDocument()

    for i, entry in enumerate(entries, 1):
        print(f"[create_word] Processing entry #{i}")

        table = doc.add_table(rows=4, cols=6)
        table.autofit = False
        widths = [Inches(1), Inches(2.5), Inches(0.8), Inches(1), Inches(0.8), Inches(1)]
        for idx, width in enumerate(widths):
            for cell in table.columns[idx].cells:
                cell.width = width

        # Prepare Title split
        title = entry["Title"]
        if len(title) > 40:
            parts = title.split(' ')
            mid = len(parts)//2
            title_text = ' '.join(parts[:mid]) + '\n' + ' '.join(parts[mid:])
        else:
            title_text = title

        # Row 1
        cells = [
            (table.cell(0,0), "Title"),
            (table.cell(0,1), title_text),
            (table.cell(0,2), "Date"),
            (table.cell(0,3), entry["Date"]),
            (table.cell(0,4), "Country"),
            (table.cell(0,5), None),  # Flag or country text filled later
        ]

        # Fill text and set vertical align center for all cells in first row
        for cell, text in cells:
            if text is not None:
                cell.text = text
            set_cell_background(cell, "ADD8E6" if text in ["Title", "Date", "Country"] else None)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Make labels bold (in label cells)
        for idx in [0,2,4]:
            cell = table.cell(0, idx)
            cell.paragraphs[0].runs[0].font.bold = True

        # Add flag or country text centered in last cell
        country_cell_value = table.cell(0,5)
        flag_path = country_flags.get(entry["Country"])
        if flag_path and os.path.isfile(flag_path):
            paragraph = country_cell_value.paragraphs[0]
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = paragraph.add_run()
            run.add_picture(flag_path, width=Inches(0.5))
            country_cell_value.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        else:
            country_cell_value.text = entry["Country"]
            country_cell_value.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Row 2 Summary
        summary_label_cell = table.cell(1,0)
        summary_label_cell.text = "Summary"
        set_cell_background(summary_label_cell, "ADD8E6")
        summary_label_cell.paragraphs[0].runs[0].font.bold = True
        summary_label_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        summary_value_cell = table.cell(1,1)
        summary_value_cell.text = entry["Summary"]
        summary_value_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for col in range(2,6):
            summary_value_cell.merge(table.cell(1,col))

        # Row 3 Link
        link_label_cell = table.cell(2,0)
        link_label_cell.text = "Link"
        set_cell_background(link_label_cell, "ADD8E6")
        link_label_cell.paragraphs[0].runs[0].font.bold = True
        link_label_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        link_value_cell = table.cell(2,1)
        link_value_cell.text = entry["Link"]
        link_value_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for col in range(2,6):
            link_value_cell.merge(table.cell(2,col))

        # Row 4 Availability
        avail_label_cell = table.cell(3,0)
        avail_label_cell.text = "Availability"
        set_cell_background(avail_label_cell, "ADD8E6")
        avail_label_cell.paragraphs[0].runs[0].font.bold = True
        avail_label_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        avail_value_cell = table.cell(3,1)
        avail_value_cell.text = entry["Availability"]
        avail_value_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for col in range(2,6):
            avail_value_cell.merge(table.cell(3,col))

        # Add page break after each table except last
        if i != len(entries):
            p = doc.add_paragraph()
            run = p.add_run()
            run.add_break(WD_BREAK.PAGE)

    doc.save(output_docx)
    print("[create_word] Document saved successfully.")


def create_pdf(entries, output_pdf):
    print("[create_pdf] Creating PDF document:", output_pdf)
    doc = SimpleDocTemplate(output_pdf, pagesize=A4,
                            rightMargin=36,leftMargin=36,
                            topMargin=36,bottomMargin=36)

    elements = []

    for i, entry in enumerate(entries, 1):
        elements.append(build_table_for_entry(entry))
        if i != len(entries):
            elements.append(Spacer(1, 0.2*inch))
            elements.append(PageBreak())

    doc.build(elements)
    print("[create_pdf] PDF saved successfully.")


def main():
    parser = argparse.ArgumentParser(description='Generate report from Word to PDF or Word.')
    parser.add_argument('input_docx', help='Input DOCX file with entries')
    parser.add_argument('-o', '--output', default='output.pdf', help='Output filename (default: output.pdf)')
    parser.add_argument('-w', '--word', action='store_true', help='Create a Word document instead of PDF')
    args = parser.parse_args()

    entries = parse_docx(args.input_docx)

    if args.word:
        # Force output to .docx extension if not provided
        if not args.output.lower().endswith('.docx'):
            output_docx = os.path.splitext(args.output)[0] + '.docx'
        else:
            output_docx = args.output
        create_word(entries, output_docx)
    else:
        # Force output to .pdf extension if not provided
        if not args.output.lower().endswith('.pdf'):
            output_pdf = os.path.splitext(args.output)[0] + '.pdf'
        else:
            output_pdf = args.output
        create_pdf(entries, output_pdf)


if __name__ == "__main__":
    main()