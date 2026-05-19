from flask import Flask, jsonify, request, Response, send_from_directory, send_file
import subprocess
import os
import io
import glob
import zipfile
import json

app = Flask(__name__)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Serve the HTML dashboard at the root URL
@app.route("/")
def dashboard():
    return send_from_directory(SCRIPT_DIR, "dashboard.html")

# ── Health check ─────────────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "API is alive"})

# ── Simple run (no streaming) – used by n8n ──────────────────────────
@app.route("/run-dhl", methods=["POST"])
def run_dhl():
    print("🚀 n8n trigger received: Running DHL Workflow...")
    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            ["python", "airtable_transform.py", "--send-to-dhl"],
            capture_output=True,
            text=True,
            cwd=SCRIPT_DIR,
            timeout=300,
            encoding="utf-8",
            errors="replace",
            env=env
        )
        return jsonify({
            "status": "completed",
            "success": result.returncode == 0,
            "output": result.stdout,
            "error": result.stderr
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ── Server-Sent Events streaming for live logs ───────────────────────
@app.route("/stream-dhl", methods=["GET"])
def stream_dhl():
    test_mode = request.args.get("test", "0") == "1"

    def generate():
        cmd = ["python", "airtable_transform.py", "--send-to-dhl"]
        if test_mode:
            cmd.append("--test")
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            cwd=SCRIPT_DIR,
            env=env,
            encoding="utf-8",
            errors="replace"
        )
        yield f"data: {json.dumps({'type': 'start', 'mode': 'test' if test_mode else 'production'})}\n\n"
        for line in iter(process.stdout.readline, ''):
            if line:
                yield f"data: {json.dumps({'type': 'log', 'line': line.rstrip()})}\n\n"
        process.wait()
        success = (process.returncode == 0)
        yield f"data: {json.dumps({'type': 'done', 'success': success})}\n\n"

    return Response(generate(), mimetype="text/event-stream")

# ── List log files ───────────────────────────────────────────────────
@app.route("/logs/list", methods=["GET"])
def list_logs():
    log_dir = os.path.join(SCRIPT_DIR, "logs")
    if not os.path.isdir(log_dir):
        return jsonify({"logs": []})
    logs = []
    for fname in os.listdir(log_dir):
        if fname.endswith(".json"):
            path = os.path.join(log_dir, fname)
            stat = os.stat(path)
            logs.append({
                "filename": fname,
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": int(stat.st_mtime)
            })
    logs.sort(key=lambda x: x["modified"], reverse=True)
    return jsonify({"logs": logs})

# ── Download latest CSV ──────────────────────────────────────────────
@app.route("/download/csv", methods=["GET"])
def download_csv():
    """Return the most recently generated DHL-Order-File CSV."""
    pattern = os.path.join(SCRIPT_DIR, "DHL-Order-File-*.csv")
    files = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)
    if not files:
        return jsonify({"error": "No CSV file found. Run a booking first."}), 404
    latest = files[0]
    return send_file(
        latest,
        as_attachment=True,
        download_name=os.path.basename(latest),
        mimetype="text/csv"
    )

# ── Download labels ZIP ──────────────────────────────────────────────
@app.route("/download/labels", methods=["GET"])
def download_labels():
    """Bundle all PDFs in the labels/ directory into a ZIP and return it."""
    labels_dir = os.path.join(SCRIPT_DIR, "labels")
    if not os.path.isdir(labels_dir):
        return jsonify({"error": "No labels directory found. Run a booking first."}), 404

    pdf_files = glob.glob(os.path.join(labels_dir, "*.pdf"))
    if not pdf_files:
        return jsonify({"error": "No label PDFs found. Run a booking first."}), 404

    # Build ZIP in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for pdf_path in sorted(pdf_files):
            zf.write(pdf_path, os.path.basename(pdf_path))
    zip_buffer.seek(0)

    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=f"DHL-Labels-{ts}.zip",
        mimetype="application/zip"
    )

# ── Download single label by order number ───────────────────────────
@app.route("/download/label/<order_number>", methods=["GET"])
def download_single_label(order_number):
    """Return a single label PDF by order number."""
    label_path = os.path.join(SCRIPT_DIR, "labels", f"label_{order_number}.pdf")
    if not os.path.exists(label_path):
        return jsonify({"error": f"Label not found for order {order_number}"}), 404
    return send_file(
        label_path,
        as_attachment=True,
        download_name=f"label_{order_number}.pdf",
        mimetype="application/pdf"
    )

# ── List available labels ────────────────────────────────────────────
@app.route("/labels/list", methods=["GET"])
def list_labels():
    """Return list of generated label PDFs."""
    labels_dir = os.path.join(SCRIPT_DIR, "labels")
    if not os.path.isdir(labels_dir):
        return jsonify({"labels": []})
    labels = []
    for fname in os.listdir(labels_dir):
        if fname.endswith(".pdf"):
            path = os.path.join(labels_dir, fname)
            stat = os.stat(path)
            labels.append({
                "filename": fname,
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": int(stat.st_mtime)
            })
    labels.sort(key=lambda x: x["modified"], reverse=True)
    return jsonify({"labels": labels, "count": len(labels)})

if __name__ == "__main__":
    print("✅ Flask API starting on port 5050...")
    app.run(host="0.0.0.0", port=5050, debug=False)