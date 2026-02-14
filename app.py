# ==== FINAL QA SYSTEM - EXACT FORMAT MATCH WITH WORKLIST MAPPING ====

from flask import Flask, render_template, request, send_file, session, jsonify, redirect, url_for
import os
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from collections import defaultdict
import pickle

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# PDF generation (using LibreOffice)
import subprocess

app = Flask(__name__)

# Secret key for session management
secret_key = os.environ.get('SECRET_KEY')
if not secret_key:
    # Generate a random key for this instance (not recommended for production)
    import secrets
    secret_key = secrets.token_hex(32)
    print("WARNING: Using randomly generated SECRET_KEY. Set SECRET_KEY environment variable in production!")

app.secret_key = secret_key

# Make sessions permanent (last 31 days)
from datetime import timedelta
app.permanent_session_lifetime = timedelta(days=31)

BASE_DIR = os.getcwd()

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
ASSET_FOLDER = os.path.join(BASE_DIR, "assets")
DATA_FOLDER = os.path.join(BASE_DIR, "data")  # For storing parsed data

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(ASSET_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

# =========================
# HELPERS
# =========================

def format_measured_value(value):
    """Convert scientific notation to 2 decimal places"""
    try:
        num = float(value)
        return f"{num:.2f}"
    except:
        return "0.00"

def format_fieldsize_mm_to_cm(field_text):
    """Convert field size from mm to cm"""
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

def format_energy_with_fff(energy, fff_value):
    """Format energy with FFF if applicable"""
    energy_str = str(energy)
    if fff_value and fff_value.upper() == "YES":
        return f"{energy_str} FFF"
    return energy_str

# =========================
# QCW PARSER - EXTRACT WORKLISTS
# =========================

def extract_worklists(file_path):
    """Extract all unique worklists from QCW file"""
    with open(file_path, 'rb') as f:
        content = f.read()
    
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    
    root = ET.fromstring(content)
    
    worklists = {}
    
    for trend in root.findall(".//TrendData"):
        worklist = trend.find("Worklist")
        if worklist is None:
            continue
            
        worklist_id = worklist.get("id")
        worklist_name_tag = worklist.find("Name")
        
        if worklist_name_tag is None or worklist_name_tag.text is None:
            worklist_name = f"Worklist_{worklist_id}"
        else:
            worklist_name = worklist_name_tag.text.strip()
        
        if worklist_id not in worklists:
            worklists[worklist_id] = worklist_name
    
    return worklists

# =========================
# QCW PARSER - FULL DATA
# =========================

def parse_qcw_with_mapping(file_path, machine_name_mapping):
    """
    Parse QCW file with custom machine name mapping.
    Organizes by date -> machine -> field size -> energies
    """
    with open(file_path, 'rb') as f:
        content = f.read()
    
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    
    root = ET.fromstring(content)
    
    # Structure: {date: {machine_name: {field_size: [energy_data]}}}
    dates_data = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    
    for trend in root.findall(".//TrendData"):
        date_str = trend.get("date").split(" ")[0]
        
        worklist = trend.find("Worklist")
        if worklist is None:
            continue
            
        worklist_id = worklist.get("id")
        
        # Get custom machine name from mapping
        machine_name = machine_name_mapping.get(worklist_id, f"Machine_{worklist_id}")
        
        # Get test details
        energy = worklist.find(".//AdminValues/Energy").text
        modality = worklist.find(".//AdminValues/Modality").text
        raw_field = worklist.find(".//AdminValues/Fieldsize").text
        field = format_fieldsize_mm_to_cm(raw_field)
        
        # Get FFF status
        fff_tag = worklist.find(".//AdminValues/FFF")
        fff_value = fff_tag.text if fff_tag is not None else "No"
        
        # Format energy with FFF
        energy_display = format_energy_with_fff(energy, fff_value)
        
        # Determine unit
        unit = energy_unit(modality)
        energy_with_unit = f"{energy_display} {unit}"
        
        # Extract measured values
        meas_data = trend.find(".//MeasData")
        if meas_data is None:
            continue
            
        analyze_values = meas_data.find("AnalyzeValues")
        if analyze_values is None:
            continue
        
        # Build measured data dictionary
        measured = {}
        for param in analyze_values:
            if param.tag == "Wedge":
                continue
            value_tag = param.find("Value")
            if value_tag is not None:
                measured[param.tag] = format_measured_value(value_tag.text)
        
        # Add to structure
        energy_data = {
            "energy": energy_with_unit,
            "measured": measured
        }
        
        dates_data[date_str][machine_name][field].append(energy_data)
    
    # Convert to regular dict and sort
    result = {}
    for date in sorted(dates_data.keys()):
        result[date] = {}
        for machine_name, fields_data in dates_data[date].items():
            result[date][machine_name] = {}
            for field_size, energy_list in fields_data.items():
                result[date][machine_name][field_size] = energy_list
    
    return result

# =========================
# DOCX GENERATION - EXACT FORMAT
# =========================

def add_custom_header_footer(doc, logo_path, institute_name):
    """Add header and footer matching the example format"""
    section = doc.sections[0]
    
    # Header with logo (right aligned)
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    if os.path.exists(logo_path):
        header_para.clear()
        run = header_para.add_run()
        run.add_picture(logo_path, width=Inches(1.2))
    
    # Footer with institute name and logo
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.clear()
    
    # Add institute name
    run = footer_para.add_run(institute_name)
    run.font.size = Pt(10)
    
    # Add small logo if exists
    if os.path.exists(logo_path):
        footer_para.add_run("   ")
        run = footer_para.add_run()
        run.add_picture(logo_path, width=Inches(0.6))
    
    footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

def add_horizontal_line(paragraph):
    """Add blue horizontal line under title"""
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')  # Line thickness
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '4472C4')  # Blue color
    
    pBdr.append(bottom)
    pPr.append(pBdr)

def generate_date_docx(date, machines_data, machine_name_mapping):
    """Generate DOCX matching exact format from example"""
    doc = Document()
    
    # Setup page
    logo_path = os.path.join(ASSET_FOLDER, "logo.png")
    name_file = os.path.join(ASSET_FOLDER, "name.txt")
    
    institute_name = "Institute/Hospital Name Here"
    if os.path.exists(name_file):
        with open(name_file) as f:
            institute_name = f.read().strip()
    
    add_custom_header_footer(doc, logo_path, institute_name)
    
    first_machine = True
    
    # Process each machine
    for machine_name, fields_data in machines_data.items():
        if not first_machine:
            doc.add_page_break()
        first_machine = False
        
        # Title: "Daily Quality Assurance"
        title_para = doc.add_heading("Daily Quality Assurance", level=1)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.runs[0]
        title_run.font.size = Pt(18)
        title_run.font.bold = True
        add_horizontal_line(title_para)
        
        doc.add_paragraph()  # Spacing
        
        # Get first energy for header (they all have same modality)
        first_field = next(iter(fields_data.values()))
        first_energy_data = first_field[0]
        first_energy = first_energy_data["energy"]
        
        # Header line: Machine Name, Date, Energy
        header_para = doc.add_paragraph()
        header_para.add_run(f"Machine Name: {machine_name}").bold = True
        header_para.add_run(f"\t\tDate: {date}").bold = True
        header_para.add_run(f"\t\tEnergy: {first_energy}").bold = True
        
        doc.add_paragraph()  # Spacing
        
        # Process each field size
        for field_size, energy_list in fields_data.items():
            # Field Size header
            field_para = doc.add_paragraph()
            field_para.add_run(f"Field Size : {field_size}").bold = True
            
            # Create table matching exact format
            # Rows: 1 header + energy rows
            num_energies = len(energy_list)
            table = doc.add_table(rows=num_energies + 1, cols=6)
            table.style = 'Table Grid'
            
            # Header row
            header_cells = table.rows[0].cells
            header_cells[0].text = "Parameter →\n\nEnergy↓"
            header_cells[1].text = "CAX"
            header_cells[2].text = "Flatness"
            header_cells[3].text = "SymmetryGT"
            header_cells[4].text = "SymmetryLR"
            header_cells[5].text = "BQF"
            
            # Make header bold and centered
            for cell in header_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Data rows
            for i, energy_data in enumerate(energy_list):
                row_cells = table.rows[i + 1].cells
                
                energy_str = energy_data["energy"]
                measured = energy_data["measured"]
                
                row_cells[0].text = energy_str
                row_cells[1].text = measured.get("CAX", "xx.xx")
                row_cells[2].text = measured.get("Flatness", "xx.xx")
                row_cells[3].text = measured.get("SymmetryGT", "xx.xx")
                row_cells[4].text = measured.get("SymmetryLR", "xx.xx")
                row_cells[5].text = measured.get("BQF", "xx.xx")
                
                # Center align data
                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Note about failed tests
            note_para = doc.add_paragraph()
            note_para.add_run("Note: 'Mention the failed tests if any and their deviation from normalized value'").italic = True
            
            doc.add_paragraph()  # Spacing between field sizes
    
    # Signature section
    doc.add_paragraph("\n\n")
    sig_para = doc.add_paragraph()
    sig_para.add_run("Signature:").bold = True
    
    # Save
    filename = f"QA_Report_{date}.docx"
    path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(path)
    return path

def convert_docx_to_pdf(docx_path):
    """Convert DOCX to PDF using LibreOffice (preserves formatting)"""
    pdf_path = docx_path.replace('.docx', '.pdf')
    
    try:
        # Try LibreOffice conversion (works on Linux/Render)
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', OUTPUT_FOLDER, docx_path
        ], check=True, capture_output=True)
        
        return pdf_path
    except:
        # Fallback: return None if conversion fails
        return None

# =========================
# CLEANUP OLD DATA FILES
# =========================

def cleanup_old_data_files():
    """Remove data files older than 24 hours"""
    try:
        current_time = datetime.now().timestamp()
        for filename in os.listdir(DATA_FOLDER):
            if filename.startswith('data_') and filename.endswith('.pkl'):
                filepath = os.path.join(DATA_FOLDER, filename)
                file_age = current_time - os.path.getmtime(filepath)
                # Delete files older than 24 hours
                if file_age > 86400:  # 24 hours in seconds
                    os.remove(filepath)
                    print(f"DEBUG: Cleaned up old data file: {filename}")
    except Exception as e:
        print(f"DEBUG: Error during cleanup: {e}")

# =========================
# ROUTES
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    # Cleanup old data files
    cleanup_old_data_files()
    
    if request.method == "POST":
        file = request.files.get("file")
        
        if not file:
            print("DEBUG: No file in request")
            return jsonify({"error": "No file uploaded"}), 400
        
        filename = file.filename
        print(f"DEBUG: File uploaded: {filename}")
        
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        print(f"DEBUG: File saved to: {file_path}")
        
        # Save institute name and logo
        inst_name = request.form.get("institute")
        if inst_name:
            with open(os.path.join(ASSET_FOLDER, "name.txt"), "w") as f:
                f.write(inst_name)
            print(f"DEBUG: Institute name saved: {inst_name}")
        
        logo = request.files.get("logo")
        if logo and logo.filename != "":
            logo.save(os.path.join(ASSET_FOLDER, "logo.png"))
            print("DEBUG: Logo saved")
        
        # Extract worklists
        print("DEBUG: Extracting worklists...")
        worklists = extract_worklists(file_path)
        print(f"DEBUG: Found {len(worklists)} worklists: {list(worklists.values())}")
        
        # Store in session (save filename, not full path)
        session.permanent = True
        session['filename'] = filename
        session['worklists'] = worklists
        session.modified = True  # Force session save
        
        print("DEBUG: Session data stored, redirecting to worklist_mapping")
        
        # Redirect to worklist mapping page
        return redirect(url_for('worklist_mapping'))
    
    return render_template("index.html")


@app.route("/worklist_mapping", methods=["GET", "POST"])
def worklist_mapping():
    """Page to map worklist IDs to custom machine names"""
    if 'worklists' not in session:
        print("DEBUG: No worklists in session on GET, redirecting to index")
        return redirect(url_for('index'))
    
    print(f"DEBUG: Worklist mapping page loaded, {len(session['worklists'])} worklists in session")
    
    if request.method == "POST":
        print("DEBUG: Worklist mapping POST received")
        
        # Get machine name mappings from form
        worklists = session.get('worklists', {})
        print(f"DEBUG: Worklists from session: {len(worklists)} items")
        
        if not worklists:
            print("DEBUG: No worklists in session, redirecting to index")
            return redirect(url_for('index'))
        
        machine_name_mapping = {}
        
        for wl_id in worklists.keys():
            custom_name = request.form.get(f"machine_{wl_id}")
            if custom_name:
                machine_name_mapping[wl_id] = custom_name.strip()
            else:
                machine_name_mapping[wl_id] = worklists[wl_id]
        
        print(f"DEBUG: Machine name mapping: {machine_name_mapping}")
        
        # Reconstruct file path from filename
        filename = session.get('filename')
        print(f"DEBUG: Filename from session: {filename}")
        
        if not filename:
            print("DEBUG: No filename in session, redirecting to index")
            return redirect(url_for('index'))
        
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        print(f"DEBUG: File path: {file_path}")
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"DEBUG: File does not exist at {file_path}, redirecting to index")
            return redirect(url_for('index'))
        
        print("DEBUG: Parsing QCW file...")
        # Parse QCW with custom names
        dates_data = parse_qcw_with_mapping(file_path, machine_name_mapping)
        print(f"DEBUG: Parsed {len(dates_data)} dates")
        
        # Save to file instead of session (session too large for 189 dates!)
        import secrets
        data_id = secrets.token_hex(16)
        data_file = os.path.join(DATA_FOLDER, f"data_{data_id}.pkl")
        
        with open(data_file, 'wb') as f:
            pickle.dump(dates_data, f)
        
        print(f"DEBUG: Saved data to file: {data_file}")
        
        # Store only the data file ID in session (small!)
        session['data_id'] = data_id
        session['machine_name_mapping'] = machine_name_mapping
        session.modified = True
        
        print("DEBUG: Redirecting to results")
        return redirect(url_for('results'))
    
    return render_template("worklist_mapping.html", worklists=session['worklists'])


@app.route("/results")
def results():
    """Display results page"""
    if 'data_id' not in session or 'machine_name_mapping' not in session:
        # Session data missing, redirect to start
        print("DEBUG: No data_id in session, redirecting to index")
        return redirect(url_for('index'))
    
    # Load data from file
    data_id = session['data_id']
    data_file = os.path.join(DATA_FOLDER, f"data_{data_id}.pkl")
    
    if not os.path.exists(data_file):
        print(f"DEBUG: Data file not found: {data_file}, redirecting to index")
        return redirect(url_for('index'))
    
    with open(data_file, 'rb') as f:
        dates_data = pickle.load(f)
    
    machine_name_mapping = session['machine_name_mapping']
    
    print(f"DEBUG: Loaded {len(dates_data)} dates from file, displaying results")
    
    return render_template("result.html", 
                         dates_data=dates_data,
                         machine_name_mapping=machine_name_mapping)


@app.route("/generate/<date>/<format>")
def generate_date_report(date, format):
    """Generate report for a specific date"""
    if 'data_id' not in session:
        return "No data available", 404
    
    # Load data from file
    data_id = session['data_id']
    data_file = os.path.join(DATA_FOLDER, f"data_{data_id}.pkl")
    
    if not os.path.exists(data_file):
        return "Data file not found", 404
    
    with open(data_file, 'rb') as f:
        dates_data = pickle.load(f)
    
    machine_name_mapping = session.get('machine_name_mapping', {})
    
    if date not in dates_data:
        return f"No data for date {date}", 404
    
    machines_data = dates_data[date]
    
    # Generate DOCX
    docx_path = generate_date_docx(date, machines_data, machine_name_mapping)
    
    if format == "pdf":
        # Convert to PDF
        pdf_path = convert_docx_to_pdf(docx_path)
        if pdf_path and os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True)
        else:
            # Fallback to DOCX if PDF conversion fails
            return send_file(docx_path, as_attachment=True)
    else:
        return send_file(docx_path, as_attachment=True)


@app.route("/generate_all/<format>")
def generate_all_dates(format):
    """Generate combined report for ALL dates"""
    if 'data_id' not in session:
        return "No data available", 404
    
    # Load data from file
    data_id = session['data_id']
    data_file = os.path.join(DATA_FOLDER, f"data_{data_id}.pkl")
    
    if not os.path.exists(data_file):
        return "Data file not found", 404
    
    with open(data_file, 'rb') as f:
        dates_data = pickle.load(f)
    
    machine_name_mapping = session.get('machine_name_mapping', {})
    
    # Create combined DOCX
    doc = Document()
    
    logo_path = os.path.join(ASSET_FOLDER, "logo.png")
    name_file = os.path.join(ASSET_FOLDER, "name.txt")
    
    institute_name = "Institute/Hospital Name Here"
    if os.path.exists(name_file):
        with open(name_file) as f:
            institute_name = f.read().strip()
    
    add_custom_header_footer(doc, logo_path, institute_name)
    
    first_page = True
    
    for date in sorted(dates_data.keys()):
        machines_data = dates_data[date]
        
        for machine_name, fields_data in machines_data.items():
            if not first_page:
                doc.add_page_break()
            first_page = False
            
            # Title
            title_para = doc.add_heading("Daily Quality Assurance", level=1)
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_para.runs[0]
            title_run.font.size = Pt(18)
            title_run.font.bold = True
            add_horizontal_line(title_para)
            
            doc.add_paragraph()
            
            # Header
            first_field = next(iter(fields_data.values()))
            first_energy_data = first_field[0]
            first_energy = first_energy_data["energy"]
            
            header_para = doc.add_paragraph()
            header_para.add_run(f"Machine Name: {machine_name}").bold = True
            header_para.add_run(f"\t\tDate: {date}").bold = True
            header_para.add_run(f"\t\tEnergy: {first_energy}").bold = True
            
            doc.add_paragraph()
            
            # Field sizes and tables
            for field_size, energy_list in fields_data.items():
                field_para = doc.add_paragraph()
                field_para.add_run(f"Field Size : {field_size}").bold = True
                
                num_energies = len(energy_list)
                table = doc.add_table(rows=num_energies + 1, cols=6)
                table.style = 'Table Grid'
                
                header_cells = table.rows[0].cells
                header_cells[0].text = "Parameter →\n\nEnergy↓"
                header_cells[1].text = "CAX"
                header_cells[2].text = "Flatness"
                header_cells[3].text = "SymmetryGT"
                header_cells[4].text = "SymmetryLR"
                header_cells[5].text = "BQF"
                
                for cell in header_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for i, energy_data in enumerate(energy_list):
                    row_cells = table.rows[i + 1].cells
                    
                    energy_str = energy_data["energy"]
                    measured = energy_data["measured"]
                    
                    row_cells[0].text = energy_str
                    row_cells[1].text = measured.get("CAX", "xx.xx")
                    row_cells[2].text = measured.get("Flatness", "xx.xx")
                    row_cells[3].text = measured.get("SymmetryGT", "xx.xx")
                    row_cells[4].text = measured.get("SymmetryLR", "xx.xx")
                    row_cells[5].text = measured.get("BQF", "xx.xx")
                    
                    for cell in row_cells:
                        for paragraph in cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                note_para = doc.add_paragraph()
                note_para.add_run("Note: 'Mention the failed tests if any and their deviation from normalized value'").italic = True
                
                doc.add_paragraph()
    
    # Signature
    doc.add_paragraph("\n\n")
    sig_para = doc.add_paragraph()
    sig_para.add_run("Signature:").bold = True
    
    filename = "QA_Report_ALL_DATES.docx"
    docx_path = os.path.join(OUTPUT_FOLDER, filename)
    doc.save(docx_path)
    
    if format == "pdf":
        pdf_path = convert_docx_to_pdf(docx_path)
        if pdf_path and os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True)
    
    return send_file(docx_path, as_attachment=True)


# =========================
# PORT FOR RENDER
# =========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
