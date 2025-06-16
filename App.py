from flask import Flask, render_template, request, send_file
import os
from ocr import mileage, MRO  # Add mro_ocr if needed

App = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        doc_type = request.form.get("doc_type")
        uploaded_file = request.files["image"]

        if uploaded_file.filename != "":
            file_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            uploaded_file.save(file_path)

            # Route to correct OCR function
            if doc_type == "mileage":
                excel_path = mileage.run(file_path, OUTPUT_FOLDER)
            elif doc_type == "MRO":
                excel_path = MRO.run(file_path, OUTPUT_FOLDER)
            else:
                return "Invalid document type selected.", 400

            return send_file(excel_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    App.run(debug=True)
