from flask import Flask, request, render_template
from docx import Document
import re
import os

app = Flask(__name__)

def parse_docx(file):
    doc = Document(file)
    data = {
        "metadata": {},
        "edits": [],
        "images": []
    }
    
    # Initialize metadata with default values
    metadata_fields = ["title", "genre", "version", "language", "secondary", "subtitles", "runtime", "date", "remarks"]
    for field in metadata_fields:
        data["metadata"][field] = ""
    
    # Parse the document
    current_section = None
    last_key = None
    for para in doc.paragraphs:
        # Replace tabs with spaces and strip leading/trailing whitespace
        text = para.text.replace('\t', ' ').strip()
        if not text:  # Skip empty lines
            continue

        # Identify sections
        if text == "Edit Notes":
            current_section = "metadata"
            continue
        elif text == "Remarks:":
            current_section = "remarks"
            continue
        elif re.match(r"\d+\.\s*\d{2}:\d{2}:\d{2}", text):
            current_section = "edits"
        elif text.endswith((".jpg", ".jpeg", ".png")):
            current_section = "images"

        # Parse based on section
        if current_section == "metadata":
            # Handle "Key: Value" format (e.g., "Version: Complete")
            if ": " in text:
                key, value = map(str.strip, text.split(": ", 1))
                if key.lower() in metadata_fields:
                    data["metadata"][key.lower()] = value
            # Handle multi-line format (e.g., "Title" followed by "Living In Two World")
            elif text in ["Title", "Genre", "Version", "Language", "Secondary", "Subtitles", "Runtime", "Date"]:
                last_key = text.lower()
            elif last_key and text:  # Value for the last key
                data["metadata"][last_key] = text
                last_key = None
        elif current_section == "remarks" and text != "Remarks:":
            data["metadata"]["remarks"] = text
        elif current_section == "edits":
            # Match edit lines like "1. 00:10:44 Edit subtitle..." or "1. 00:02:58 - 00:03:15 Edit out..."
            match = re.match(r"\d+\.\s*(\d{2}:\d{2}:\d{2}(?:\s*-\s*\d{2}:\d{2}:\d{2})?)\s*(.+)", text)
            if match:
                time, description = match.groups()
                data["edits"].append({"time": time, "description": description.strip()})
        elif current_section == "images":
            if text.endswith((".jpg", ".jpeg", ".png")):
                data["images"].append(text)

    # Set filename and default remarks
    data["metadata"]["filename"] = os.path.basename(file.filename)
    if not data["metadata"]["remarks"]:
        data["metadata"]["remarks"] = "None"
    
    return data

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename.endswith(".docx"):
            data = parse_docx(file)
            return render_template("index.html", data=data)
        return render_template("index.html", error="Please upload a .docx file")
    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
