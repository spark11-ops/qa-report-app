# ==== IMPROVED QA SYSTEM - DATE-BASED GROUPING WITH DYNAMIC WORKLISTS ====

from flask import Flask, render_template, request, send_file, session, jsonify
import os
import xml.etree.ElementTree as ET
from datetime import datetime
from collections import defaultdict

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-to-random-secret-key-in-production')

BASE_DIR = os.getcwd()

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
ASSET_FOLDER = os.path.join(BASE_DIR, "assets")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(ASSET_FOLDER, exist_ok=True)

# =========================
# HELPERS
# =========================

def format_decimal(value):
    try:
        return f"{float(value):.4f}"
    except:
        return value

def check_pass_fail(value, min_val, max_val):
    try:
        v = float(value)
        return "PASS" if float(min_val) <= v <= float(max_val) else "FAIL"
    except:
        return "NA"

def calc_deviation(value, target):
    try:
        v = float(value)
        t = float(target)
        if t == 0:
            return ""
        return f"{((v - t) / t) * 100:.2f} %"
    except:
        return ""

def format_fieldsize_mm_to_cm(field_text):
    try:
        x_mm, y_mm = field_text.split("x")
        x_cm = float(x_mm) / 10
        y_cm = float(y_mm) / 10
        return f"{x_cm:.0f} cm X {y_cm:.0f} cm"
    except:
        return field_text

def energy_unit(modality):
    """Determine energy unit based on modality"""
    return "MV" if modality.lower().startswith("photon") else "MeV"

# =========================
# QCW PARSER (DATE-BASED)
# =========================

def parse_qcw(file_path):
    """
    Parse QCW file and organize data by DATE.
    Returns: {
        'date1': [
            {
                'worklist_name': 'Unique New',
                'worklist_id': '1234',
                'entries': [
                    {
                        'energy': '6',
                        'field': '10 cm X 10 cm',
                        'modality': 'Photons',
                        'data': [parameter dicts...]
                    },
                    ...
                ]
            },
            ...
        ],
        'date2': [...],
        ...
    }
    """
    
    # Read file with BOM handling
    with open(file_path, 'rb') as f:
        content = f.read()
    
    # Remove BOM if present
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    
    root = ET.fromstring(content)
    
    # Organize by date
    dates_data = defaultdict(lambda: defaultdict(lambda: {'entries': []}))
    
    for trend in root.findall(".//TrendData"):
        # Extract date (YYYY-MM-DD format)
        date_str = trend.get("date").split(" ")[0]
        
        # Get worklist info
        worklist = trend.find("Worklist")
        if worklist is None:
            continue
            
        worklist_id = worklist.get("id")
        worklist_name_tag = worklist.find("Name")
        
        if worklist_name_tag is None or worklist_name_tag.text is None:
            worklist_name = f"Worklist_{worklist_id}"
        else:
            worklist_name = worklist_name_tag.text.strip()
        
        # Get measurement data
        energy = worklist.find(".//AdminValues/Energy").text
        modality = worklist.find(".//AdminValues/Modality").text
        raw_field = worklist.find(".//AdminValues/Fieldsize").text
        field = format_fieldsize_mm_to_cm(raw_field)
        
        # Get parameters
        analyze_values = worklist.findall(".//MeasData/AnalyzeValues/*")
        analyze_params = worklist.findall(".//AdminData/AnalyzeParams/*")
        
        tolerance_map = {}
        target_map = {}
        
        for p in analyze_params:
            if p.tag == "Wedge":
                continue
            tolerance_map[p.tag] = (p.find("Min").text, p.find("Max").text)
            target_map[p.tag] = p.find("Target").text
        
        parameters = []
        
        for val in analyze_values:
            if val.tag == "Wedge":
                continue
            
            name = val.tag
            value = val.find("Value").text
            
            min_val, max_val = tolerance_map.get(name, ("", ""))
            target = target_map.get(name, "")
            
            parameters.append({
                "name": name,
                "value": format_decimal(value),
                "target": format_decimal(target),
                "tolerance": f"{format_decimal(min_val)} to {format_decimal(max_val)}",
                "deviation": calc_deviation(value, target),
                "status": check_pass_fail(value, min_val, max_val)
            })
        
        # Create entry for this measurement
        entry = {
            "energy": energy,
            "field": field,
            "modality": modality,
            "data": parameters
        }
        
        # Add to date -> worklist structure
        if worklist_id not in dates_data[date_str]:
            dates_data[date_str][worklist_id] = {
                'worklist_name': worklist_name,
                'worklist_id': worklist_id,
                'entries': []
            }
        
        dates_data[date_str][worklist_id]['entries'].append(entry)
    
    # Convert to regular dict and sort dates
    result = {}
    for date in sorted(dates_data.keys()):
        result[date] = list(dates_data[date].values())
    
    return result

# =========================
# ANALYTICS
# =========================

def calculate_analytics(dates_data):
    """Calculate summary statistics"""
    analytics = {
        "total_tests": 0,
        "total_passes": 0,
        "total_fails": 0,
        "pass_rate": 0,
        "total_dates": len(dates_data),
        "total_worklists": set(),
        "by_date": {}
    }
    
    for date, worklists in dates_data.items():
        date_stats = {
            "total_tests": 0,
            "passes": 0,
            "fails": 0,
            "worklists": len(worklists)
        }
        
        for worklist in worklists:
            analytics["total_worklists"].add(worklist['worklist_name'])
            
            for entry in worklist['entries']:
                for param in entry["data"]:
                    analytics["total_tests"] += 1
                    date_stats["total_tests"] += 1
                    
                    if param["status"] == "PASS":
                        analytics["total_passes"] += 1
                        date_stats["passes"] += 1
                    elif param["status"] == "FAIL":
                        analytics["total_fails"] += 1
                        date_stats["fails"] += 1
        
        if date_stats["total_tests"] > 0:
            date_stats["pass_rate"] = (date_stats["passes"] / date_stats["total_tests"]) * 100
        else:
            date_stats["pass_rate"] = 0
            
        analytics["by_date"][date] = date_stats
    
    if analytics["total_tests"] > 0:
        analytics["pass_rate"] = (analytics["total_passes"] / analytics["total_tests"]) * 100
    
    analytics["total_worklists"] = len(analytics["total_worklists"])
    
    return analytics

# =========================
# HEADER + FOOTER
# =========================

def add_header_footer(doc):
    section = doc.sections[0]
    logo_path = os.path.join(ASSET_FOLDER, "logo.png")
    name_file = os.path.join(ASSET_FOLDER, "name.txt")

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    if os.path.exists(logo_path):
        hp.clear()
        hp.add_run().add_picture(logo_path, width=Inches(1.2))

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.clear()
    fp.add_run("\n" + "_"*90 + "\n")

    institute = "Institute Name"
    if os.path.exists(name_file):
        with open(name_file) as f:
            institute = f.read().strip()

    fp.add_run(institute)

    if os.path.exists(logo_path):
        fp.add_run("   ")
        fp.add_run().add_picture(logo_path, width=Inches(0.8))

# =========================
# DOCX GENERATION (DATE-BASED)
# =========================

def generate_date_docx(date, worklists_data):
    """Generate a single DOCX for a specific date with all worklists"""
    doc = Document()
    add_header_footer(doc)

    # Add title
    title = doc.add_heading("Daily Quality Assurance Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add date info
    date_para = doc.add_paragraph()
    date_run = date_para.add_run(f"Date: {date}")
    date_run.bold = True
    date_run.font.size = Pt(14)
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()  # Spacing

    # Process each worklist
    for i, worklist in enumerate(worklists_data):
        if i > 0:
            doc.add_page_break()

        worklist_name = worklist['worklist_name']

        # Worklist header
        wl_heading = doc.add_heading(f"Machine: {worklist_name}", level=1)
        wl_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Process each entry (different energies/field sizes)
        for j, entry in enumerate(worklist['entries']):
            if j > 0:
                doc.add_paragraph()  # Spacing between entries
                doc.add_paragraph("─" * 80)
                doc.add_paragraph()

            # Entry details
            doc.add_paragraph(f"Type of Radiation: {entry['modality']}")
            
            unit = energy_unit(entry["modality"])
            doc.add_paragraph(f"Energy: {entry['energy']} {unit}")
            
            doc.add_paragraph(f"Field Size: {entry['field']}")
            
            doc.add_paragraph()  # Spacing

            # Parameters table
            table = doc.add_table(rows=1, cols=6)
            table.style = "Table Grid"

            headers = ["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]
            for k, h in enumerate(headers):
                table.rows[0].cells[k].text = h

            for row in entry["data"]:
                cells = table.add_row().cells
                cells[0].text = row["name"]
                cells[1].text = row["value"]
                cells[2].text = row["target"]
                cells[3].text = row["tolerance"]
                cells[4].text = row["deviation"]
                cells[5].text = row["status"]

    # Signature section
    doc.add_paragraph("\n\n\n\n\n\n\n\nSignature: _______________________")

    # Save
    path = os.path.join(OUTPUT_FOLDER, f"QA_Report_{date}.docx")
    doc.save(path)
    return path

# =========================
# PDF GENERATION (DATE-BASED)
# =========================

def generate_date_pdf(date, worklists_data):
    """Generate a single PDF for a specific date with all worklists"""
    path = os.path.join(OUTPUT_FOLDER, f"QA_Report_{date}.pdf")
    styles = getSampleStyleSheet()

    doc_pdf = SimpleDocTemplate(path, pagesize=A4)
    elements = []

    # Title
    elements.append(Paragraph("Daily Quality Assurance Report", styles['Title']))
    elements.append(Paragraph(f"Date: {date}", styles['Heading2']))
    elements.append(Spacer(1, 20))

    # Process each worklist
    for worklist in worklists_data:
        worklist_name = worklist['worklist_name']
        
        elements.append(Paragraph(f"Machine: {worklist_name}", styles['Heading1']))
        elements.append(Spacer(1, 10))

        # Process each entry
        for entry in worklist['entries']:
            elements.append(Paragraph(f"Type of Radiation: {entry['modality']}", styles['Normal']))
            
            unit = energy_unit(entry["modality"])
            elements.append(Paragraph(f"Energy: {entry['energy']} {unit}", styles['Normal']))
            
            elements.append(Paragraph(f"Field Size: {entry['field']}", styles['Normal']))
            elements.append(Spacer(1, 10))

            # Parameters table
            data = [["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]]

            for row in entry["data"]:
                data.append([
                    row["name"],
                    row["value"],
                    row["target"],
                    row["tolerance"],
                    row["deviation"],
                    row["status"]
                ])

            table = Table(data)
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey)
            ]))

            elements.append(table)
            elements.append(Spacer(1, 20))

        elements.append(PageBreak())

    doc_pdf.build(elements)
    return path

# =========================
# ROUTES
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        
        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        inst_name = request.form.get("institute")
        if inst_name:
            with open(os.path.join(ASSET_FOLDER, "name.txt"), "w") as f:
                f.write(inst_name)

        logo = request.files.get("logo")
        if logo and logo.filename != "":
            logo.save(os.path.join(ASSET_FOLDER, "logo.png"))

        # Parse QCW file (organized by date)
        dates_data = parse_qcw(file_path)
        session['dates_data'] = dates_data
        session['analytics'] = calculate_analytics(dates_data)
        session['filename'] = file.filename
        
        return render_template("result.html", 
                             dates_data=dates_data,
                             analytics=session['analytics'])

    return render_template("index.html")


@app.route("/api/analytics")
def get_analytics():
    """API endpoint for fetching analytics data"""
    if 'analytics' in session:
        return jsonify(session['analytics'])
    return jsonify({"error": "No data available"}), 404


@app.route("/generate/<date>/<format>")
def generate_date_report(date, format):
    """Generate report for a specific date"""
    
    # Use cached data from session
    if 'dates_data' not in session:
        # Fallback to file parsing
        files = os.listdir(UPLOAD_FOLDER)
        if not files:
            return "No file available", 404
        file_path = os.path.join(UPLOAD_FOLDER, files[0])
        dates_data = parse_qcw(file_path)
    else:
        dates_data = session['dates_data']
    
    if date not in dates_data:
        return f"No data for date {date}", 404
    
    worklists_data = dates_data[date]

    if format == "docx":
        path = generate_date_docx(date, worklists_data)
    else:  # pdf
        path = generate_date_pdf(date, worklists_data)

    return send_file(path, as_attachment=True)


@app.route("/generate_all/<format>")
def generate_all_dates(format):
    """Generate combined report for ALL dates"""
    
    if 'dates_data' not in session:
        files = os.listdir(UPLOAD_FOLDER)
        if not files:
            return "No file available", 404
        file_path = os.path.join(UPLOAD_FOLDER, files[0])
        dates_data = parse_qcw(file_path)
    else:
        dates_data = session['dates_data']

    if format == "docx":
        # Create combined DOCX
        doc = Document()
        add_header_footer(doc)

        for i, (date, worklists_data) in enumerate(sorted(dates_data.items())):
            if i > 0:
                doc.add_page_break()

            # Date-based content
            title = doc.add_heading("Daily Quality Assurance Report", 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            date_para = doc.add_paragraph()
            date_run = date_para.add_run(f"Date: {date}")
            date_run.bold = True
            date_run.font.size = Pt(14)
            date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph()

            for j, worklist in enumerate(worklists_data):
                if j > 0:
                    doc.add_page_break()

                worklist_name = worklist['worklist_name']
                wl_heading = doc.add_heading(f"Machine: {worklist_name}", level=1)

                for k, entry in enumerate(worklist['entries']):
                    if k > 0:
                        doc.add_paragraph()
                        doc.add_paragraph("─" * 80)
                        doc.add_paragraph()

                    doc.add_paragraph(f"Type of Radiation: {entry['modality']}")
                    unit = energy_unit(entry["modality"])
                    doc.add_paragraph(f"Energy: {entry['energy']} {unit}")
                    doc.add_paragraph(f"Field Size: {entry['field']}")
                    doc.add_paragraph()

                    table = doc.add_table(rows=1, cols=6)
                    table.style = "Table Grid"

                    headers = ["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]
                    for l, h in enumerate(headers):
                        table.rows[0].cells[l].text = h

                    for row in entry["data"]:
                        cells = table.add_row().cells
                        cells[0].text = row["name"]
                        cells[1].text = row["value"]
                        cells[2].text = row["target"]
                        cells[3].text = row["tolerance"]
                        cells[4].text = row["deviation"]
                        cells[5].text = row["status"]

            doc.add_paragraph("\n\n\n\n\n\n\n\nSignature: _______________________")

        path = os.path.join(OUTPUT_FOLDER, "QA_Report_ALL_DATES.docx")
        doc.save(path)

    else:  # pdf
        path = os.path.join(OUTPUT_FOLDER, "QA_Report_ALL_DATES.pdf")
        styles = getSampleStyleSheet()
        doc_pdf = SimpleDocTemplate(path, pagesize=A4)
        elements = []

        for date, worklists_data in sorted(dates_data.items()):
            elements.append(Paragraph("Daily Quality Assurance Report", styles['Title']))
            elements.append(Paragraph(f"Date: {date}", styles['Heading2']))
            elements.append(Spacer(1, 20))

            for worklist in worklists_data:
                worklist_name = worklist['worklist_name']
                elements.append(Paragraph(f"Machine: {worklist_name}", styles['Heading1']))
                elements.append(Spacer(1, 10))

                for entry in worklist['entries']:
                    elements.append(Paragraph(f"Type of Radiation: {entry['modality']}", styles['Normal']))
                    unit = energy_unit(entry["modality"])
                    elements.append(Paragraph(f"Energy: {entry['energy']} {unit}", styles['Normal']))
                    elements.append(Paragraph(f"Field Size: {entry['field']}", styles['Normal']))
                    elements.append(Spacer(1, 10))

                    data = [["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]]
                    for row in entry["data"]:
                        data.append([
                            row["name"],
                            row["value"],
                            row["target"],
                            row["tolerance"],
                            row["deviation"],
                            row["status"]
                        ])

                    table = Table(data)
                    table.setStyle(TableStyle([
                        ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey)
                    ]))

                    elements.append(table)
                    elements.append(Spacer(1, 20))

            elements.append(PageBreak())

        doc_pdf.build(elements)

    return send_file(path, as_attachment=True)


# =========================
# PORT FOR RENDER
# =========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
