from flask import Flask, render_template, request, send_file
import os
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import subprocess

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ASSET_FOLDER = "assets"

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

def convert_docx_to_pdf(docx_path):
    output_dir = os.path.dirname(docx_path)

    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        docx_path,
        "--outdir", output_dir
    ], check=True)

    return docx_path.replace(".docx", ".pdf")

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
# REMOVE BORDERS (HEADER ROW)
# =========================

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement('w:tblBorders')

    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'nil')
        borders.append(element)

    tblPr.append(borders)

    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            tcBorders = OxmlElement('w:tcBorders')
            for edge in ('top', 'left', 'bottom', 'right'):
                e = OxmlElement(f'w:{edge}')
                e.set(qn('w:val'), 'nil')
                tcBorders.append(e)

            tcPr.append(tcBorders)

# =========================
# QA TABLE STYLE
# =========================

def style_table(table):
    table.style = "Table Grid"

# =========================
# HEADER + FOOTER
# =========================

def add_header_footer(doc):
    section = doc.sections[0]

    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    logo_path = os.path.join(ASSET_FOLDER, "logo.png")
    if os.path.exists(logo_path):
        p.add_run().add_picture(logo_path, width=Inches(1.0))

    footer = section.footer
    p = footer.paragraphs[0]

    p.add_run("\n" + "_" * 90 + "\n")

    institute_name = "Your Institute"
    name_file = os.path.join(ASSET_FOLDER, "name.txt")
    if os.path.exists(name_file):
        institute_name = open(name_file).read().strip()

    p.add_run(institute_name)

    if os.path.exists(logo_path):
        p.add_run().add_picture(logo_path, width=Inches(0.6))

# =========================
# DOCX GENERATION (single)
# =========================

def generate_docx(machine, entry):
    doc = Document()
    add_header_footer(doc)

    title = doc.add_heading("Daily Quality Assurance", 0)
    title.alignment = 1

    info_table = doc.add_table(rows=1, cols=2)
    info_table.autofit = True
    remove_table_borders(info_table)

    left = info_table.rows[0].cells[0]
    right = info_table.rows[0].cells[1]

    left.text = f"Machine Name: {machine}"
    right.text = f"Date: {entry['date']}"
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph(f"Type of Radiation : {entry['modality']}")
    doc.add_paragraph(f"Energy : {entry['energy']} MV")
    doc.add_paragraph(f"Field Size : {entry['field']}")

    table = doc.add_table(rows=1, cols=6)
    style_table(table)

    headers = ["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True

    for row in entry["data"]:
        cells = table.add_row().cells
        cells[0].text = row["name"]
        cells[1].text = row["value"]
        cells[2].text = row["target"]
        cells[3].text = row["tolerance"]
        cells[4].text = row["deviation"]
        cells[5].text = row["status"]

    doc.add_paragraph("\nSignature:")

    filename = f"{machine}_{entry['date']}_{entry['energy']}_{entry['field']}.docx"
    path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(path)
    return path

# =========================
# DOCX GENERATION (combined)
# =========================

def generate_combined_docx(machine, entries):
    doc = Document()
    add_header_footer(doc)

    for i, entry in enumerate(entries):

        title = doc.add_heading("Daily Quality Assurance", 0)
        title.alignment = 1

        info_table = doc.add_table(rows=1, cols=2)
        info_table.autofit = True
        remove_table_borders(info_table)

        left = info_table.rows[0].cells[0]
        right = info_table.rows[0].cells[1]

        left.text = f"Machine Name: {machine}"
        right.text = f"Date: {entry['date']}"
        right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_paragraph(f"Type of Radiation : {entry['modality']}")
        doc.add_paragraph(f"Energy : {entry['energy']} MV")
        doc.add_paragraph(f"Field Size : {entry['field']}")

        table = doc.add_table(rows=1, cols=6)
        style_table(table)

        headers = ["Parameter", "Measured", "Target", "Tolerance", "Deviation", "Status"]

        for j, h in enumerate(headers):
            cell = table.rows[0].cells[j]
            cell.text = h
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True

        for row in entry["data"]:
            cells = table.add_row().cells
            cells[0].text = row["name"]
            cells[1].text = row["value"]
            cells[2].text = row["target"]
            cells[3].text = row["tolerance"]
            cells[4].text = row["deviation"]
            cells[5].text = row["status"]

        doc.add_paragraph("\nSignature:")

        if i != len(entries) - 1:
            doc.add_page_break()

    filename = f"{machine}_ALL_QA.docx"
    path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(path)
    return path

# =========================
# ROUTES
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        file = request.files.get("file")
        if file:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)

        inst_name = request.form.get("institute")
        if inst_name:
            with open(os.path.join(ASSET_FOLDER, "name.txt"), "w") as f:
                f.write(inst_name)

        logo = request.files.get("logo")
        if logo and logo.filename != "":
            logo.save(os.path.join(ASSET_FOLDER, "logo.png"))

        machines = parse_qcw(file_path)
        return render_template("result.html", machines=machines)

    return render_template("index.html")


@app.route("/generate/<machine>/<int:index>/<format>")
def generate(machine, index, format):
    file_path = os.path.join(UPLOAD_FOLDER, os.listdir(UPLOAD_FOLDER)[0])
    machines = parse_qcw(file_path)

    entry = machines[machine][index]

    docx_path = generate_docx(machine, entry)

    if format == "docx":
        return send_file(docx_path, as_attachment=True)

    pdf_path = convert_docx_to_pdf(docx_path)
    return send_file(pdf_path, as_attachment=True)


@app.route("/generate_all/<machine>/docx")
def generate_all_docx(machine):
    file_path = os.path.join(UPLOAD_FOLDER, os.listdir(UPLOAD_FOLDER)[0])
    machines = parse_qcw(file_path)

    docx_path = generate_combined_docx(machine, machines[machine])
    return send_file(docx_path, as_attachment=True)


@app.route("/generate_all/<machine>/pdf")
def generate_all_pdf(machine):
    file_path = os.path.join(UPLOAD_FOLDER, os.listdir(UPLOAD_FOLDER)[0])
    machines = parse_qcw(file_path)

    docx_path = generate_combined_docx(machine, machines[machine])
    pdf_path = convert_docx_to_pdf(docx_path)

    return send_file(pdf_path, as_attachment=True)


if __name__ == "__main__":
    app.run()
