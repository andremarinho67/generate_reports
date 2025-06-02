from docx import Document as DocxDocument
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak
import os
import re

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


def create_pdf(entries, output_pdf):
    print("[create_pdf] Creating PDF:", output_pdf)
    doc = SimpleDocTemplate(output_pdf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []

    for i, entry in enumerate(entries, 1):
        print(f"[create_pdf] Processing entry #{i}")
        table = build_table_for_entry(entry)
        elements.append(table)

        # Add page break after each table except the last one
        if i < len(entries):
            elements.append(PageBreak())

    doc.build(elements)
    print("[create_pdf] PDF creation done.")

if __name__ == "__main__":
    input_docx = "input.docx"
    output_pdf = "output.pdf"

    print("[main] Starting script...")
    entries = parse_docx(input_docx)
    print(f"[main] Parsed {len(entries)} entries from document.")

    create_pdf(entries, output_pdf)
    print(f"[main] PDF generated: {output_pdf}")
