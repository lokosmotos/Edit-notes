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

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    """Check if the file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def parse_docx(file_path):
    """Parse a .docx file into metadata, edits, and images."""
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

    current_section = "metadata"
    last_metadata_key = None

    for para in doc.paragraphs:
        # Normalize text: replace tabs and multiple spaces with a single space
        text = re.sub(r'\s+', ' ', para.text).strip()
        if not text:
            continue

        # Section detection
        if text.lower() == "edit notes":
            current_section = "metadata"
            continue
        elif text.lower() == "remarks:":
            current_section = "remarks"
            continue
        elif re.match(r"^\d+\.\s*\d{2}:\d{2}:\d{2}", text):
            current_section = "edits"
            continue
        elif text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
            current_section = "images"

        # Parse metadata
        if current_section == "metadata":
            if text in ["Title", "Genre", "Version", "Language", "Secondary", "Subtitles", "Runtime", "Date"]:
                last_metadata_key = text.lower()
            elif last_metadata_key and text:
                if ": " in text:  # Handle "Version: Complete"
                    key, value = map(str.strip, text.split(": ", 1))
                    if key.lower() in data["metadata"]:
                        data["metadata"][key.lower()] = value
                else:
                    data["metadata"][last_metadata_key] = text
                    last_metadata_key = None

        # Parse remarks
        elif current_section == "remarks":
            if text != "Remarks:":
                data["metadata"]["remarks"] = text or "None"

        # Parse edits
        elif current_section == "edits":
            match = re.match(r"(\d+)\.\s*(\d{2}:\d{2}:\d{2}(?:\s*-\s*\d{2}:\d{2}:\d{2})?)\s*(.+)", text)
            if match:
                edit_number, timestamp, description = match.groups()
                data["edits"].append({
                    "number": int(edit_number),
                    "time": timestamp,
                    "description": description.strip()
                })

        # Parse images
        elif current_section == "images":
            if text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                data["images"].append(text)

    return data

@app.route('/', methods=['GET', 'POST'])
def index():
    """Handle file uploads and render the main page."""
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
    """Export parsed data as a JSON file."""
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
    """Serve a file for download."""
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
