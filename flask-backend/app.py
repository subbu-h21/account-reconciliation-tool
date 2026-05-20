from dotenv import load_dotenv
load_dotenv()

import json
import queue
import tempfile
import threading
import uuid
import time
import os
import logging
from flask import (Flask, render_template, request, send_file, jsonify,
                   after_this_request, make_response, Response, stream_with_context)
from werkzeug.utils import secure_filename
from tasks import process_files_task

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ALLOWED_EXTENSIONS = {'.xls', '.xlsx'}

# session_id -> {"queue": Queue|None, "output_path": str, "result": dict|None, "done": threading.Event}
active_streams = {}


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def save_upload(file):
    _, ext = os.path.splitext(secure_filename(file.filename))
    fd, path = tempfile.mkstemp(prefix="upload_", suffix=ext)
    os.close(fd)
    file.save(path)
    return path


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process_files():
    if 'ac_statement' not in request.files or 'our_books' not in request.files:
        return jsonify({"error": "Both files are required"}), 400

    ac_file = request.files['ac_statement']
    ob_file = request.files['our_books']

    if not allowed_file(ac_file.filename) or not allowed_file(ob_file.filename):
        return jsonify({"error": "Only .xls and .xlsx files are allowed"}), 400

    use_ai      = request.form.get('ai_matching') == 'on'
    ai_provider = request.form.get('ai_provider', 'gemini')

    if use_ai:
        if ai_provider == 'gemini' and not os.getenv("GEMINI_API_KEY"):
            return jsonify({"error": "AI matching requested but GEMINI_API_KEY is not set"}), 400
        if ai_provider == 'openai' and not (os.getenv("OPEN_ROUTER_API_KEY") or "").strip():
            return jsonify({"error": "AI matching requested but OPEN_ROUTER_API_KEY is not set"}), 400

    ac_path = save_upload(ac_file)
    ob_path = save_upload(ob_file)

    timestamp  = int(time.time())
    session_id = uuid.uuid4().hex
    out = os.path.join(
        app.config['UPLOAD_FOLDER'],
        f"combined_{timestamp}_{session_id}.xlsx"
    )

    log_queue  = queue.Queue() if use_ai else None
    done_event = threading.Event()

    active_streams[session_id] = {
        "queue":       log_queue,
        "output_path": out,
        "result":      None,
        "done":        done_event,
    }

    def run():
        try:
            result = process_files_task(
                ac_path, ob_path, out, timestamp,
                use_ai=use_ai,
                ai_provider=ai_provider,
                log_queue=log_queue,
            )
        except Exception as e:
            logger.exception("Background task crashed: %s", e)
            result = {"path": out, "ai_ran": False}
        active_streams[session_id]["result"] = result
        done_event.set()
        if log_queue:
            log_queue.put(None)  # SSE sentinel

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"session_id": session_id, "use_ai": use_ai})


@app.route('/stream/<session_id>')
def stream(session_id):
    session = active_streams.get(session_id)
    if not session or not session["queue"]:
        return jsonify({"error": "Stream not found"}), 404

    def generate():
        q = session["queue"]
        deadline = time.time() + 600
        while True:
            try:
                msg = q.get(timeout=15)
            except queue.Empty:
                if time.time() > deadline:
                    payload = json.dumps({"ok": False, "error": "stream timeout — model took too long"})
                    yield f"event: done\ndata: {payload}\n\n"
                    return
                # SSE comment — keeps the connection alive without triggering a message event
                yield ": keepalive\n\n"
                continue
            if msg is None:
                result  = session.get("result") or {}
                ai_ran  = "ran" if result.get("ai_ran") else "skipped"
                ok      = os.path.exists(session["output_path"])
                payload = json.dumps({"ok": ok, "ai_ran": ai_ran})
                yield f"event: done\ndata: {payload}\n\n"
                return
            yield f"event: log\ndata: {msg}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route('/result/<session_id>')
def get_result(session_id):
    session = active_streams.get(session_id)
    if not session:
        return jsonify({"error": "Session not found or already downloaded"}), 404

    if not session["done"].wait(timeout=600):
        return jsonify({"error": "Processing timed out"}), 504

    active_streams.pop(session_id, None)

    out = session["output_path"]
    if not os.path.exists(out):
        return jsonify({"error": "Processing failed — check server logs"}), 500

    result = session.get("result") or {}
    ai_ran = "ran" if result.get("ai_ran") else "not-ran"

    @after_this_request
    def cleanup(response):
        try:
            os.remove(out)
        except Exception:
            pass
        return response

    resp = make_response(send_file(out, as_attachment=True, download_name="reconciliation.xlsx"))
    resp.headers['X-AI-Ran'] = ai_ran
    return resp


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
