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
    
    # Parse metadata
    current_section = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if text == "Edit Notes":
            current_section = "metadata"
        elif text == "Remarks:":
            current_section = "remarks"
        elif text.startswith("1."):
            current_section = "edits"
        elif text.endswith((".jpg", ".jpeg", ".png")):
            current_section = "images"
        
        if current_section == "metadata" and ":" in text:
            key, value = map(str.strip, text.split(":", 1))
            if key in ["Title", "Genre", "Version", "Language", "Secondary", "Subtitles", "Runtime", "Date"]:
                data["metadata"][key.lower()] = value
        elif current_section == "remarks" and text and text != "Remarks:":
            data["metadata"]["remarks"] = text
        elif current_section == "edits" and text and re.match(r"\d+\.\s+\d{2}:\d{2}:\d{2}", text):
            time_match = re.match(r"\d+\.\s+(\d{2}:\d{2}:\d{2}(?:\s*-\s*\d{2}:\d{2}:\d{2})?)\s+(.+)", text)
            if time_match:
                time, description = time_match.groups()
                data["edits"].append({"time": time, "description": description})
        elif current_section == "images" and text and text.endswith((".jpg", ".jpeg", ".png")):
            data["images"].append(text)
    
    data["metadata"]["filename"] = os.path.basename(file.filename)
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
