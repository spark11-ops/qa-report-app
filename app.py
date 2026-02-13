from flask import Flask, render_template, request, send_file
import os
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

# ===== Render persistent storage =====
BASE_DIR = "/data"

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

def normalize_machine(name):
    if "Unique" in name:
        return "Unique"
    if "TrueBeam" in name:
        return "TrueBeam"
    return name

def format_fieldsize_mm_to_cm(field_text):
    try:
        x_mm, y_mm = field_text.split("x")
        x_cm = float(x_mm) / 10
        y_cm = float(y_mm) / 10
        return f"{x_cm:.0f} cm X {y_cm:.0f} cm"
    except:
        return field_text

# =========================
# QCW PARSER
# =========================

def parse_qcw(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    machines = {}

    for trend in root.findall(".//TrendData"):
        raw_machine = trend.find(".//Worklist/Name").text
        machine = normalize_machine(raw_machine)
        date = trend.get("date").split(" ")[0]

        energy = trend.find(".//AdminValues/Energy").text
        modality = trend.find(".//AdminValues/Modality").text
        raw_field = trend.find(".//AdminValues/Fieldsize").text
        field = format_fieldsize_mm_to_cm(raw_field)

        analyze_values = trend.findall(".//MeasData/AnalyzeValues/*")
        analyze_params = trend.findall(".//AdminData/AnalyzeParams/*")

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

        entry = {
            "date": date,
            "energy": energy,
            "field": field,
            "modality": modality,
            "data": parameters
        }

        machines.setdefault(machine, []).append(entry)

    return machines

# =========================
# DOCX GENERATION
# =========================

def generate_docx(machine, entry):
    doc = Document()

    doc.add_heading("Daily Quality Assurance", 0)

    doc.add_paragraph(f"Machine Name: {machine}")
    doc.add_paragraph(f"Date: {entry['date']}")
    doc.add_paragraph(f"Type of Radiation: {entry['modality']}")
    doc.add_paragraph(f"Energy: {entry['energy']} MV")
    doc.add_paragraph(f"Field Size: {entry['field']}")

    table = doc.add_table(rows=1, cols=6)
    headers = ["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]

    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h

    for row in entry["data"]:
        cells = table.add_row().cells
        cells[0].text = row["name"]
        cells[1].text = row["value"]
        cells[2].text = row["target"]
        cells[3].text = row["tolerance"]
        cells[4].text = row["deviation"]
        cells[5].text = row["status"]

    filename = f"{machine}_{entry['date']}.docx"
    path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(path)
    return path

# =========================
# PDF GENERATION (REPORTLAB)
# =========================

def generate_pdf(machine, entry):
    filename = f"{machine}_{entry['date']}.pdf"
    path = os.path.join(OUTPUT_FOLDER, filename)

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(path, pagesize=A4)

    elements = []

    elements.append(Paragraph("Daily Quality Assurance", styles['Title']))
    elements.append(Spacer(1, 10))

    elements.append(Paragraph(f"Machine: {machine}", styles['Normal']))
    elements.append(Paragraph(f"Date: {entry['date']}", styles['Normal']))
    elements.append(Paragraph(f"Type of Radiation: {entry['modality']}", styles['Normal']))
    elements.append(Paragraph(f"Energy: {entry['energy']} MV", styles['Normal']))
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
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey)
    ]))

    elements.append(table)

    doc.build(elements)
    return path

# =========================
# ROUTES
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        machines = parse_qcw(file_path)
        return render_template("result.html", machines=machines)

    return render_template("index.html")


@app.route("/generate/<machine>/<int:index>/<format>")
def generate(machine, index, format):
    file_path = os.path.join(UPLOAD_FOLDER, os.listdir(UPLOAD_FOLDER)[0])
    machines = parse_qcw(file_path)
    entry = machines[machine][index]

    if format == "docx":
        path = generate_docx(machine, entry)
    else:
        path = generate_pdf(machine, entry)

    return send_file(path, as_attachment=True)


@app.route("/generate_all/<machine>/<format>")
def generate_all(machine, format):
    file_path = os.path.join(UPLOAD_FOLDER, os.listdir(UPLOAD_FOLDER)[0])
    machines = parse_qcw(file_path)

    paths = []
    for entry in machines[machine]:
        if format == "docx":
            paths.append(generate_docx(machine, entry))
        else:
            paths.append(generate_pdf(machine, entry))

    # Return first file for now (simple version)
    return send_file(paths[0], as_attachment=True)


if __name__ == "__main__":
    app.run()
