from flask import Flask, request, render_template, jsonify, send_from_directory
from docx import Document
import re
import os
from werkzeug.utils import secure_filename
import json
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'docx'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def parse_docx(file_path):
    doc = Document(file_path)
    data = {
        "metadata": {
            "title": "",
            "genre": "",
            "version": "",
            "language": "",
            "secondary": "",
            "subtitles": "",
            "runtime": "",
            "date": "",
            "remarks": "None",
            "filename": os.path.basename(file_path),
            "parsed_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "edits": [],
        "images": []
    }
    
    current_section = None
    last_key = None
    remarks_content = []
    
    for para in doc.paragraphs:
        text = para.text.replace('\t', ' ').strip()
        if not text:
            continue

        # Section detection
        if text.lower() == "edit notes":
            current_section = "metadata"
            continue
        elif text.lower().startswith("remarks"):
            current_section = "remarks"
            continue
        elif re.match(r"^\d+\.\s*\d{2}:\d{2}:\d{2}", text):
            current_section = "edits"
        elif text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
            current_section = "images"

        # Section parsing
        if current_section == "metadata":
            if ": " in text:
                key, value = map(str.strip, text.split(": ", 1))
                key_lower = key.lower()
                if key_lower in data["metadata"]:
                    data["metadata"][key_lower] = value
            elif text.lower() in [k.lower() for k in data["metadata"].keys()]:
                last_key = text.lower()
            elif last_key:
                data["metadata"][last_key] = text
                last_key = None
        
        elif current_section == "remarks":
            if not text.lower().startswith("remarks"):
                remarks_content.append(text)
        
        elif current_section == "edits":
            match = re.match(r"^(\d+)\.\s*(\d{2}:\d{2}:\d{2}(?:\s*-\s*\d{2}:\d{2}:\d{2})?)\s*(.+)", text)
            if match:
                number, time, description = match.groups()
                data["edits"].append({
                    "number": int(number),
                    "time": time,
                    "description": description.strip()
                })
        
        elif current_section == "images":
            if text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                data["images"].append(text)

    if remarks_content:
        data["metadata"]["remarks"] = "\n".join(remarks_content)
    
    return data

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='No file selected')
            
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No file selected')
            
        if not allowed_file(file.filename):
            return render_template('index.html', error='Only .docx files are allowed')
            
        try:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            data = parse_docx(filepath)
            return render_template('index.html', data=data)
            
        except Exception as e:
            return render_template('index.html', error=f'Error processing file: {str(e)}')
    
    return render_template('index.html')

@app.route('/export/json', methods=['POST'])
def export_json():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data provided'}), 400
    
    filename = secure_filename(f"{data['metadata']['filename']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.json")
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=2)
    
    return jsonify({
        'download_url': f'/download/{filename}',
        'filename': filename
    })

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
