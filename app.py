from flask import Flask, render_template, request, send_file
import os
from translit import convert_docx
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_file = request.files["file"]
        if uploaded_file and uploaded_file.filename.endswith(".docx"):
            filename = secure_filename(uploaded_file.filename)
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            name, ext = os.path.splitext(filename)
            output_filename = f"converted_{name}.docx"
            output_path = os.path.join(CONVERTED_FOLDER, output_filename)
            uploaded_file.save(input_path)
            convert_docx(input_path, output_path)
            return send_file(output_path, as_attachment=True, download_name=output_filename)
        else:
            return "Invalid file. Please upload a .docx file."
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
