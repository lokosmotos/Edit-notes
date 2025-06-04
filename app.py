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

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def parse_docx(file_path):
    doc = Document(file_path)
    data = {
        "metadata": {
            "title": "N/A",
            "genre": "N/A",
            "version": "N/A",
            "language": "N/A",
            "secondary": "N/A",
            "subtitles": "N/A",
            "runtime": "N/A",
            "date": "N/A",
            "remarks": "None",
            "filename": os.path.basename(file_path),
            "parsed_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "edits": [],
        "images": []
    }

    current_section = None
    collecting_remarks = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Section detection
        if text.lower() == "edit notes":
            current_section = "metadata"
            continue
        elif text.lower().startswith("remarks"):
            current_section = "remarks"
            if ":" in text:
                data["metadata"]["remarks"] = text.split(":", 1)[1].strip()
            collecting_remarks = True
            continue
        elif re.match(r"^\d+\.\s+(\d{2}:\d{2}:\d{2})(?:\s*-\s*(\d{2}:\d{2}:\d{2}))?", text):
            current_section = "edits"
            collecting_remarks = False
        elif text.lower().endswith((".jpg", ".jpeg", ".png")):
            current_section = "images"
            collecting_remarks = False

        # Metadata parsing
        if current_section == "metadata":
            if ": " in text:
                key, value = [part.strip() for part in text.split(": ", 1)]
                key_lower = key.lower()
                if key_lower in data["metadata"]:
                    data["metadata"][key_lower] = value
            elif ":" in text:  # Handle cases like "Remarks: None"
                key, value = [part.strip() for part in text.split(":", 1)]
                key_lower = key.lower()
                if key_lower in data["metadata"]:
                    data["metadata"][key_lower] = value

        # Edit parsing
        elif current_section == "edits":
            # Match formats:
            # 1. 00:02:58 - 00:03:15 Edit out bikini scene
            # 2. 00:06:57 Edit subtitle: "dammit" (oh God)
            match = re.match(r"^(\d+)\.\s+(\d{2}:\d{2}:\d{2})(?:\s*-\s*(\d{2}:\d{2}:\d{2}))?\s*(.+)", text)
            if match:
                number = match.group(1)
                start_time = match.group(2)
                end_time = match.group(3)
                description = match.group(4).strip()
                
                time_display = start_time
                if end_time:
                    time_display += f" - {end_time}"
                
                data["edits"].append({
                    "number": int(number),
                    "time": time_display,
                    "description": description
                })

        # Remarks collection
        elif collecting_remarks:
            if not text.lower().startswith("remarks"):
                if data["metadata"]["remarks"] == "None":
                    data["metadata"]["remarks"] = text
                else:
                    data["metadata"]["remarks"] += "\n" + text

        # Image collection
        elif current_section == "images":
            if text.lower().endswith((".jpg", ".jpeg", ".png")):
                data["images"].append(text)

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
    
    return render_template('index.html', data=None)

@app.route('/export/json', methods=['POST'])
def export_json():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        filename = secure_filename(f"{data.get('metadata', {}).get('filename', 'edit_notes')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.json")
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        with open(filepath, 'w') as f:
            json.dump(data, f, indent=2)
        
        return jsonify({
            'download_url': f'/download/{filename}',
            'filename': filename
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
