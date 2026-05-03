import os
import uuid
import base64
import json
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
from utils.docx_handler import extract_images_docx, replace_images_docx
from utils.pptx_handler import extract_images_pptx, replace_images_pptx
from utils.xlsx_handler import extract_images_xlsx, replace_images_xlsx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = "temp_uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"docx", "pptx", "xlsx", "xlsm", "xltx", "potx"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_handler(ext):
    ext = ext.lower()
    if ext == "docx":
        return extract_images_docx, replace_images_docx
    if ext in ("pptx", "potx"):
        return extract_images_pptx, replace_images_pptx
    if ext in ("xlsx", "xlsm", "xltx"):
        return extract_images_xlsx, replace_images_xlsx
    return None, None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/extract", methods=["POST"])
def extract():
    file = request.files.get("file")
    if not file or not allowed_file(file.filename):
        return jsonify({"error": "Ugyldig filtype"}), 400

    ext = file.filename.rsplit(".", 1)[1].lower()
    session_id = str(uuid.uuid4())
    save_path = os.path.join(UPLOAD_FOLDER, f"{session_id}.{ext}")
    file.save(save_path)

    extract_fn, _ = get_handler(ext)
    if not extract_fn:
        return jsonify({"error": "Ikke støttet filtype"}), 400

    images = extract_fn(save_path)
    return jsonify({"session_id": session_id, "ext": ext, "images": images})


@app.route("/replace", methods=["POST"])
def replace():
    data = request.form
    session_id = data.get("session_id")
    ext = data.get("ext")
    selected = json.loads(data.get("selected", "[]"))

    logo_file = request.files.get("logo")
    if not logo_file:
        return jsonify({"error": "Ingen logo lastet opp"}), 400

    logo_bytes = logo_file.read()
    logo_ext = logo_file.filename.rsplit(".", 1)[-1].lower() if "." in logo_file.filename else "png"

    src_path = os.path.join(UPLOAD_FOLDER, f"{session_id}.{ext}")
    if not os.path.exists(src_path):
        return jsonify({"error": "Økt ikke funnet — last opp filen på nytt"}), 400

    out_id = str(uuid.uuid4())
    out_path = os.path.join(UPLOAD_FOLDER, f"{out_id}_rebranded.{ext}")

    _, replace_fn = get_handler(ext)
    replace_fn(src_path, out_path, selected, logo_bytes, logo_ext)

    return send_file(
        out_path,
        as_attachment=True,
        download_name=f"rebranded_{session_id[:8]}.{ext}",
    )


@app.route("/cleanup", methods=["POST"])
def cleanup():
    session_id = request.json.get("session_id", "")
    for f in os.listdir(UPLOAD_FOLDER):
        if f.startswith(session_id):
            try:
                os.remove(os.path.join(UPLOAD_FOLDER, f))
            except Exception:
                pass
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
