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
    
    current_section = "metadata"  # Start in metadata section
    last_key = None
    
    for i, para in enumerate(doc.paragraphs):
        # Normalize text: replace multiple spaces/tabs with a single space, strip whitespace
        text = re.sub(r'\s+', ' ', para.text).strip()
        if not text:
            continue

        # Section detection
        if text.lower() == "edit notes":
            current_section = "metadata"
            continue
        elif text.lower().startswith("remarks"):
            current_section = "remarks"
            if ":" in text:
                data["metadata"]["remarks"] = text.split(":", 1)[1].strip() or "None"
            continue
        elif re.match(r"^\d+\.\s*\d{2}:\d{2}:\d{2}", text):
            current_section = "edits"
        elif text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")) and i > len(doc.paragraphs) * 0.9:  # Last 10% of document
            current_section = "images"

        # Parse metadata
        if current_section == "metadata":
            if text in ["Title", "Genre", "Version", "Language", "Secondary", "Subtitles", "Runtime", "Date"]:
                last_key = text.lower()
            elif ": " in text:  # Handle "Version: Complete"
                key, value = map(str.strip, text.split(": ", 1))
                if key.lower() in data["metadata"]:
                    data["metadata"][key.lower()] = value
            elif last_key and text and not any(k.lower() in text.lower() for k in ["Title", "Genre", "Version", "Language", "Secondary", "Subtitles", "Runtime", "Date"]):
                data["metadata"][last_key] = text
                last_key = None

        # Parse remarks
        elif current_section == "remarks" and not text.lower().startswith("remarks"):
            data["metadata"]["remarks"] = text.strip() or "None"
            
        # Parse edits
        elif current_section == "edits":
            # Match lines like "1. 00:10:44 Edit subtitle..." or "1. 00:02:58 - 00:03:15 Edit out..."
            match = re.match(r"(\d+)\.\s*(\d{2}:\d{2}:\d{2}(?:\s*-\s*\d{2}:\d{2}:\d{2})?)\s*(.+)", text)
            if match:
                edit_number, timestamp, description = match.groups()
                try:
                    edit_number = int(edit_number.strip('.'))
                    data["edits"].append({
                        "number": edit_number,
                        "time": timestamp,
                        "description": description.strip()
                    })
                except (ValueError, IndexError) as e:
                    print(f"Error parsing edit line: {text}\nError: {e}")
                    continue

        # Parse images
        elif current_section == "images":
            if text.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
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
    
    return render_template('index.html', data=None)  # Explicitly pass None for initial load

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
