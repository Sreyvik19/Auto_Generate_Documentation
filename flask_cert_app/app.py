from flask import Flask, render_template, request, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime
import csv
import sys
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)

# ---------------- Configuration ----------------
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///certs.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
UPLOAD_FILE = os.path.join(STATIC_DIR, 'certificates')
TEMPLATE_PATH = os.path.join(STATIC_DIR, 'certificate.png')  # <-- make sure this exists

# Ensure folders exist
os.makedirs(UPLOAD_FILE, exist_ok=True)
app.config['UPLOAD_FILE'] = UPLOAD_FILE

db = SQLAlchemy(app)

# ---------------- Database Model ----------------
class StudentCertificate(db.Model):
    __tablename__ = 'student_certificates'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    student_class = db.Column(db.String(50), nullable=False)
    certificate = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

with app.app_context():
    db.create_all()

# ---------------- Certificate Generation ----------------
def generate_certificate_image(name, student_class):
    # Check if template exists
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Certificate template not found at {TEMPLATE_PATH}")

    img = Image.open(TEMPLATE_PATH).convert("RGBA")
    draw = ImageDraw.Draw(img)

    # Use TTF font if available
    font_path = "arial.ttf"
    try:
        name_font = ImageFont.truetype(font_path, 50)
        class_font = ImageFont.truetype(font_path, 30)
        date_font = ImageFont.truetype(font_path, 25)
    except Exception:
        name_font = ImageFont.load_default()
        class_font = ImageFont.load_default()
        date_font = ImageFont.load_default()

    # Coordinates for text (adjust to your template)
    name_position = (600, 400)
    class_position = (600, 500)
    date_position = (600, 600)

    draw.text(name_position, name, fill="black", font=name_font)
    draw.text(class_position, f"Class: {student_class}", fill="black", font=class_font)
    draw.text(date_position, f"Date: {datetime.now().strftime('%Y-%m-%d')}", fill="black", font=date_font)

    # Create a safe, unique filename
    safe_name = secure_filename(name) or "student"
    timestamp = int(datetime.utcnow().timestamp())
    filename = f"{safe_name}_{timestamp}_certificate.png"
    filepath = os.path.join(UPLOAD_FILE, filename)

    img.save(filepath)
    return filename

# ---------------- Routes ----------------
@app.route('/')
def home():
    return redirect(url_for('generate_page'))

@app.route('/generate')
def generate_page():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_certificate():
    # Handle CSV upload (process file in-memory; do not save uploaded CSV to a folder)
    csv_file = request.files.get('csv_file')
    if csv_file and csv_file.filename and secure_filename(csv_file.filename).lower().endswith('.csv'):
        # Read uploaded file content in-memory
        raw = csv_file.read()
        if isinstance(raw, bytes):
            try:
                text = raw.decode('utf-8')
            except Exception:
                text = raw.decode('latin-1')
        else:
            text = str(raw)

        reader = csv.DictReader(io.StringIO(text))
        for row in reader:
            name = (row.get('name') or row.get('Name') or "").strip()
            student_class = (row.get('class') or row.get('Class') or "").strip()
            if not name or not student_class:
                continue
            filename = generate_certificate_image(name, student_class)
            cert = StudentCertificate(name=name, student_class=student_class, certificate=filename)
            db.session.add(cert)
        db.session.commit()
        return redirect(url_for('student_list'))

    # Single student form (fallback)
    name = request.form.get('name', '').strip()
    student_class = request.form.get('student_class', '').strip()
    if not name or not student_class:
        return "Name and Class are required", 400

    filename = generate_certificate_image(name, student_class)
    cert = StudentCertificate(name=name, student_class=student_class, certificate=filename)
    db.session.add(cert)
    db.session.commit()

    return redirect(url_for('student_list'))

@app.route('/students')
def student_list():
    students = StudentCertificate.query.order_by(StudentCertificate.created_at.desc()).all()
    return render_template('students.html', students=students)

if __name__ == '__main__':
    # Extra safety check
    if not os.path.exists(TEMPLATE_PATH):
        print(f"ERROR: Certificate template not found at {TEMPLATE_PATH}", file=sys.stderr)
        print("Make sure 'certificate.png' is in the 'static' folder.", file=sys.stderr)
        sys.exit(1)

    app.run(debug=True)
