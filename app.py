import os
import sys
import io
import uuid
import json
import copy
import threading
import contextlib
import builtins
import subprocess
import time
import re
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_file, session
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "workspace"
UPLOAD_DIR = WORK_DIR / "uploads"
OUTPUT_DIR = WORK_DIR / "outputs"
PREVIEW_DIR = WORK_DIR / "previews"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
PREVIEW_DIR.mkdir(parents=True, exist_ok=True)

# Reuse the mature document engine for DOCX parsing, formatting, tables,
# images, undo/redo, autosave, AI editing, document Q&A and summarization.
# The web layer adds multi-document workspace state and batch commands.
sys.path.insert(0, str(BASE_DIR))
import document_editor_core as core

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-in-production")
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024  # 250 MB safety limit per request
@app.errorhandler(413)
def request_too_large(_error):
    return jsonify({"error": "Upload is too large. The maximum combined upload size is 250 MB."}), 413


EDITOR_LOCK = threading.RLock()


def _sid():
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    return session["sid"]


def _session_dir():
    path = WORK_DIR / "sessions" / _sid()
    path.mkdir(parents=True, exist_ok=True)
    return path


def _bind_core(state):
    """Bind the active in-memory document state to the legacy core engine."""
    core.current_doc = state.get("doc")
    core.original_doc = state.get("original")
    core.current_doc_path = state.get("path")
    core.AUTOSAVE_CURRENT_PATH = state.get("autosave_path")
    core.UNDO_STACK = state.get("undo", [])
    core.REDO_STACK = state.get("redo", [])
    core.AUTOSAVE_DIR = str(_session_dir() / "autosave")
    core.IMAGES_DIR = str(_session_dir() / "images")
    core.TEMPLATES_DIR = str(BASE_DIR / "templates")
    core.LEGAL_CHECK_DIR = str(_session_dir() / "legal_checks")
    core.AUTOSAVE_ENABLED = state.get("autosave_enabled", True)
    core.ensure_working_folders()


def _capture(fn):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        result = fn()
    return result, buf.getvalue().strip()


def _state_from_core():
    return {
        "doc": core.current_doc,
        "original": core.original_doc,
        "path": core.current_doc_path,
        "autosave_path": core.AUTOSAVE_CURRENT_PATH,
        "undo": core.UNDO_STACK,
        "redo": core.REDO_STACK,
        "autosave_enabled": core.AUTOSAVE_ENABLED,
    }


app_state = {}


def _empty_state():
    return {"documents": {}, "active": None}


def _get_state():
    return app_state.setdefault(_sid(), _empty_state())


def _active_state(state):
    key = state.get("active")
    return state.get("documents", {}).get(key) if key else None


def _sync_active(state):
    active = _active_state(state)
    if active is not None:
        _bind_core(active)
    else:
        _bind_core({"doc": None, "original": None, "path": None, "undo": [], "redo": [],
                    "autosave_path": None, "autosave_enabled": True})


def _store_active(state):
    """Persist all mutable core globals back into the active document record."""
    if state.get("active") is not None:
        state.setdefault("documents", {})[state["active"]] = _state_from_core()


def _document_records(state):
    records = []
    for key, doc_state in state.get("documents", {}).items():
        path = doc_state.get("path") or key
        records.append({
            "id": key,
            "name": os.path.basename(path),
            "active": key == state.get("active"),
            "stats": _doc_stats(doc_state.get("doc"), doc_state.get("path")),
        })
    return records


def _activate(state, key):
    if key not in state.get("documents", {}):
        return False
    _store_active(state)
    state["active"] = key
    _sync_active(state)
    return True


def _unique_document_id(state, filename):
    stem = Path(filename).stem
    key = stem
    n = 2
    while key in state.get("documents", {}):
        key = f"{stem}-{n}"
        n += 1
    return key


def _doc_stats(doc, path=None):
    if doc is None:
        return None
    images = 0
    for p in doc.paragraphs:
        if "<w:drawing" in p._element.xml or "drawing" in p._element.xml:
            images += 1
    return {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "sections": len(doc.sections),
        "images": images,
        "name": os.path.basename(path or (core.current_doc_path if doc is core.current_doc else "document.docx")),
    }


def _preview(doc, limit=80):
    if doc is None:
        return []
    items = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        items.append({
            "index": i,
            "type": core.identify_paragraph_type(p),
            "text": text,
        })
        if len(items) >= limit:
            break
    return items


def _make_pdf_preview():
    """Render the current DOCX through LibreOffice for a faithful, full-document preview."""
    if core.current_doc is None:
        return None

    session_preview = _session_dir() / "preview"
    session_preview.mkdir(parents=True, exist_ok=True)
    docx_path = session_preview / "working_document.docx"
    pdf_path = session_preview / "working_document.pdf"
    profile = session_preview / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)

    # Save the in-memory working copy without changing the editor's active path.
    core.current_doc.save(str(docx_path))
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            pass

    soffice = "libreoffice"
    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(session_preview),
        f"-env:UserInstallation=file://{profile}",
        str(docx_path),
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
    except (FileNotFoundError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError("LibreOffice is required to render the document preview.") from exc

    # LibreOffice names the output after the input file.
    generated = session_preview / "working_document.pdf"
    if result.returncode != 0 or not generated.exists():
        detail = (result.stderr or result.stdout or "unknown LibreOffice error").strip()
        raise RuntimeError(f"Could not render document preview: {detail[-500:]}")
    return generated


def _preview_url():
    return f"/api/preview.pdf?rev={time.time_ns()}"


def _looks_done(text):
    normalized = " ".join(text.lower().strip().split())
    return normalized in {
        "done", "i'm done", "im done", "finished", "finish", "that's all",
        "thats all", "download", "download it", "give me the document",
        "give me the edited document", "export", "export it"
    }


@app.get("/")
def index():
    state = _get_state()
    with EDITOR_LOCK:
        _sync_active(state)
        stats = _doc_stats(core.current_doc)
        preview = _preview(core.current_doc)
        documents = _document_records(state)
    return render_template("index.html", stats=stats, preview=preview, documents=documents)


@app.post("/api/upload")
def upload():
    files = request.files.getlist("files") or request.files.getlist("file")
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify({"error": "Please choose at least one document."}), 400

    allowed = {".docx", ".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp"}
    state = _get_state()
    added, auxiliary, errors = [], [], []

    with EDITOR_LOCK:
        for file in files:
            suffix = Path(file.filename).suffix.lower()
            if suffix not in allowed:
                errors.append(f"{file.filename}: unsupported file type")
                continue
            # Keep individual uploads reasonably bounded while allowing a large
            # multi-file workspace upload.
            try:
                file.stream.seek(0, os.SEEK_END)
                file_size = file.stream.tell()
                file.stream.seek(0)
            except (OSError, AttributeError):
                file_size = 0
            if file_size <= 0:
                errors.append(f"{file.filename}: file is empty")
                continue
            if file_size > 50 * 1024 * 1024:
                errors.append(f"{file.filename}: file is larger than 50 MB")
                continue
            name = secure_filename(file.filename) or f"upload{suffix or '.docx'}"
            # Avoid overwriting another upload with the same filename.
            path = _session_dir() / "uploads" / name
            path.parent.mkdir(parents=True, exist_ok=True)
            if path.exists():
                base, ext = path.stem, path.suffix
                n = 2
                while path.exists():
                    path = path.parent / f"{base}-{n}{ext}"
                    n += 1
                name = path.name
            file.save(path)

            if suffix != ".docx":
                auxiliary.append({"name": name, "type": suffix[1:].upper(), "path": str(path)})
                continue

            try:
                doc = core.Document(str(path))
                import copy as _copy
                key = _unique_document_id(state, name)
                state["documents"][key] = {
                    "doc": doc,
                    "original": _copy.deepcopy(doc),
                    "path": str(path),
                    "autosave_path": None,
                    "undo": [],
                    "redo": [],
                    "autosave_enabled": True,
                }
                added.append(key)
            except Exception as exc:
                errors.append(f"{name}: {exc}")

        if added:
            # New uploads become the active selection; previously loaded files remain available.
            state["active"] = added[-1]
            _sync_active(state)
            core.autosave_current_document("initial load")
            _store_active(state)

        return jsonify({
            "ok": bool(added or auxiliary),
            "added": added,
            "auxiliary": auxiliary,
            "errors": errors,
            "active": state.get("active"),
            "documents": _document_records(state),
            "stats": _doc_stats(core.current_doc),
            "preview": _preview(core.current_doc),
            "preview_url": _preview_url() if core.current_doc is not None else None,
            "message": f"Loaded {len(added)} DOCX file(s)." if added else "Files uploaded.",
        })


@app.get("/api/state")
def api_state():
    with EDITOR_LOCK:
        state = _get_state()
        _sync_active(state)
        return jsonify({
            "loaded": core.current_doc is not None,
            "active": state.get("active"),
            "documents": _document_records(state),
            "stats": _doc_stats(core.current_doc),
            "preview": _preview(core.current_doc),
            "preview_url": _preview_url() if core.current_doc is not None else None,
            "can_undo": bool(core.UNDO_STACK),
            "can_redo": bool(core.REDO_STACK),
        })


@app.post("/api/select")
def select_document():
    payload = request.get_json(silent=True) or {}
    key = str(payload.get("id", "")).strip()
    with EDITOR_LOCK:
        state = _get_state()
        if not _activate(state, key):
            return jsonify({"error": "Document is not loaded."}), 404
        _store_active(state)
        return jsonify({
            "ok": True,
            "active": key,
            "documents": _document_records(state),
            "stats": _doc_stats(core.current_doc),
            "preview": _preview(core.current_doc),
            "preview_url": _preview_url(),
            "can_undo": bool(core.UNDO_STACK),
            "can_redo": bool(core.REDO_STACK),
        })


@app.get("/api/preview.pdf")
def preview_pdf():
    with EDITOR_LOCK:
        state = _get_state()
        _sync_active(state)
        if core.current_doc is None:
            return jsonify({"error": "Upload a DOCX document first."}), 404
        try:
            pdf = _make_pdf_preview()
        except RuntimeError as exc:
            return jsonify({"error": str(exc)}), 500
        return send_file(pdf, mimetype="application/pdf", max_age=0)


@contextlib.contextmanager
def _web_command_environment():
    """Make terminal commands operate on files uploaded to this web session."""
    upload_root = _session_dir() / "uploads"
    upload_root.mkdir(parents=True, exist_ok=True)
    old_cwd = os.getcwd()
    old_input = builtins.input
    try:
        os.chdir(upload_root)
        # A few terminal commands ask for y/N confirmation. In the web UI the
        # user's explicit command is treated as approval so the request never
        # blocks waiting for terminal input.
        builtins.input = lambda prompt="": "y"
        yield upload_root
    finally:
        builtins.input = old_input
        os.chdir(old_cwd)


def _session_file_snapshot(root):
    snap = {}
    try:
        for path in root.iterdir():
            if path.is_file() and path.suffix.lower() in {".docx", ".pdf"}:
                try:
                    snap[path.name] = path.stat().st_mtime_ns
                except OSError:
                    pass
    except OSError:
        pass
    return snap


def _session_file_links(root, before=None):
    links = []
    before = before or {}
    try:
        for path in sorted(root.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True):
            if path.is_file() and path.suffix.lower() in {".docx", ".pdf"}:
                try:
                    if path.name in before and path.stat().st_mtime_ns == before[path.name]:
                        continue
                except OSError:
                    continue
                links.append({"name": path.name, "url": f"/api/session-file/{path.name}"})
    except OSError:
        pass
    return links


def _command_is_mutating(text):
    low = text.lower().strip()
    return (
        low.startswith("doc edit") or low.startswith("table edit") or
        low.startswith("edit table ") or low.startswith("replace image ") or
        low.startswith("change ") or low.startswith("replace ") or
        low.startswith("rewrite paragraph") or low.startswith("rewrite paragraphs") or
        low.startswith("simplify paragraph") or low.startswith("simplify paragraphs") or
        low in ("reset", "improve format", "improve formatting", "format") or
        low.startswith("format --") or low.startswith("reconstruct ") or
        low.startswith("improvements --apply")
    )


def _run_single_command(state, text):
    """Run one command against the currently active document and persist it."""
    _sync_active(state)

    # `load <file>` / `open <file>` selects an already-uploaded document instead
    # of creating a second hidden core document.
    m = re.match(r"^(?:load|open)\s+(.+)$", text.strip(), re.I)
    if m:
        wanted = m.group(1).strip().strip('"')
        for key, doc_state in state.get("documents", {}).items():
            if os.path.basename(doc_state.get("path", "")).lower() == os.path.basename(wanted).lower() or key.lower() == wanted.lower():
                _activate(state, key)
                _store_active(state)
                return {
                    "output": f"[{os.path.basename(core.current_doc_path)}] Loaded.",
                    "changed": False,
                    "files": [],
                }
        return {"output": f"Couldn't find '{wanted}' among loaded documents. Use 'list docs'.",
                "changed": False, "files": []}

    with _web_command_environment() as upload_root:
        file_snapshot = _session_file_snapshot(upload_root)
        before = copy.deepcopy(core.current_doc) if core.current_doc is not None else None
        before_xml = before._element.xml if before is not None else None
        capture = io.StringIO()
        with contextlib.redirect_stdout(capture):
            if text.lower().strip() == "exit":
                print("The web session stays open. Use the browser to continue working.")
            elif text.startswith("--"):
                question = text[2:].strip()
                if question:
                    core.chat_with_openrouter(question, max_tokens=500)
            elif core.handle_convert_command(text):
                pass
            elif core.handle_doc_command(text):
                pass
            elif core.current_doc is not None and core._natural_prompt_is_likely_edit(text):
                before_natural = copy.deepcopy(core.current_doc)
                _, natural_output = _capture(lambda: core.natural_document_edit(text))
                if natural_output:
                    print(natural_output)
            else:
                return None

        if core.current_doc_path and not os.path.isabs(core.current_doc_path):
            core.current_doc_path = str((upload_root / core.current_doc_path).resolve())

        after_xml = core.current_doc._element.xml if core.current_doc is not None else None
        changed = before_xml != after_xml
        low = text.lower().strip()
        history_command = (
            _command_is_mutating(text)
            or (core.current_doc is not None and core._natural_prompt_is_likely_edit(text))
        ) and low not in {"undo", "u", "redo", "r"}
        if changed and before is not None and core.current_doc is not None and history_command:
            core._push_undo_snapshot(before)
            core.autosave_current_document("web command")
        _store_active(state)
        output = capture.getvalue().strip()
        files = _session_file_links(upload_root, file_snapshot)
        return {"output": output or "Command completed.", "changed": changed, "files": files}


def _run_terminal_command(state, text):
    """Run a command, optionally across every loaded DOCX with `--all`."""
    raw = text.strip()
    if not raw:
        return None

    # --all is a command modifier, not part of the underlying command syntax.
    # It can appear at the end: change "Load" to "Road" --all
    # or before a command: --all change "Load" to "Road"
    all_mode = bool(re.search(r"\s--all\s*$", raw, re.I) or re.match(r"^--all(?:\s+|$)", raw, re.I))
    if not all_mode:
        return _run_single_command(state, raw)

    command = re.sub(r"\s--all\s*$", "", raw, flags=re.I).strip()
    command = re.sub(r"^--all\s+", "", command, flags=re.I).strip()
    if not command:
        return {"output": "Usage: <command> --all  (example: change \"Load\" to \"Road\" --all)",
                "changed": False, "files": []}

    documents = list(state.get("documents", {}).keys())
    original_active = state.get("active")
    if not documents:
        return {"output": "No DOCX documents are loaded.", "changed": False, "files": []}

    reports = []
    changed_any = False
    all_files = []
    for key in documents:
        state["active"] = key
        result = _run_single_command(state, command)
        if result is None:
            reports.append(f"[{key}] Unknown command: {command}")
            continue
        changed_any |= bool(result["changed"])
        all_files.extend(result.get("files", []))
        reports.append(f"[{os.path.basename(state['documents'][key].get('path',''))}] {result['output'] or 'completed'}")

    # Restore the user's original active document after a batch operation.
    state["active"] = original_active if original_active in documents else documents[-1]
    _sync_active(state)
    _store_active(state)
    return {
        "output": "Applied to all loaded documents.\n\n" + "\n".join(reports),
        "changed": changed_any,
        "files": all_files,
    }


@app.post("/api/chat")
def chat():
    payload = request.get_json(silent=True) or {}
    text = str(payload.get("message", "")).strip()
    if not text:
        return jsonify({"error": "Message cannot be empty."}), 400

    with EDITOR_LOCK:
        state = _get_state()
        _sync_active(state)

        if core.current_doc is None and not text.lower().strip() in {"help", "commands", "list docs"} and not text.startswith("--"):
            # Some commands such as list docs/help can be useful before loading;
            # all other document commands naturally report that no document is loaded.
            pass

        # First give the exact terminal command dispatcher a chance. This makes
        # the web chat a command console as well as a natural-language assistant.
        command_result = _run_terminal_command(state, text)
        if command_result is not None:
            response = command_result["output"]
            files = command_result["files"]
            data = {
                "ok": True,
                "type": "command",
                "message": response,
                "changed": command_result["changed"],
                "stats": _doc_stats(core.current_doc),
                "preview_url": _preview_url() if core.current_doc is not None else None,
                "can_undo": bool(core.UNDO_STACK),
                "can_redo": bool(core.REDO_STACK),
                "files": files,
            }
            # Loading a different document via `open` / `load` must update the
            # web session state immediately.
            return jsonify(data)

        if core.current_doc is None:
            if text.startswith("--"):
                question = text[2:].strip()
                if not question:
                    return jsonify({"error": "Usage: -- <general AI question>"}), 400
                answer = core.chat_with_openrouter(question, max_tokens=700, stream_to_screen=False).strip()
                return jsonify({"ok": True, "type": "answer", "message": answer or "I couldn't produce an answer."})
            return jsonify({"error": "Upload a DOCX document first, then enter a command or natural-language request."}), 400

        # Terminal behaviour for natural-language input:
        # likely edits mutate the document; plain questions use document Q&A.
        before = copy.deepcopy(core.current_doc)
        before_xml = before._element.xml
        if core._natural_prompt_is_likely_edit(text):
            _, output = _capture(lambda: core.natural_document_edit(text))
            changed = before_xml != core.current_doc._element.xml
            if changed:
                core._push_undo_snapshot(before)
                core.autosave_current_document("web natural edit")
            response_text = output or ("Applied the requested document change." if changed else
                                       "I couldn't find a safe matching change to apply.")
            kind = "edit"
        else:
            context = core.extract_document_text(core.current_doc, max_chars=60000)
            prompt = (
                f"DOCUMENT: {core.document_name()}\n\n"
                f"DOCUMENT CONTENT:\n{context}\n\n"
                f"QUESTION: {text}\n\n"
                "Answer the user's question using only the supplied document. "
                "If the document does not contain enough information, say so. "
                "Be concise but useful. Mention paragraph/table locations when useful."
            )
            response_text = core.chat_with_openrouter(
                prompt,
                system_prompt=(
                    "You are the AI assistant inside a professional document editor. "
                    "Answer questions strictly from the supplied document. Never invent facts. "
                    "Return only the answer for the user, without terminal labels or prefixes."
                ),
                max_tokens=900,
                stream_to_screen=False,
            ).strip()
            if not response_text:
                response_text = "I couldn't produce an answer from the document."
            changed = False
            kind = "answer"

        _store_active(state)
        return jsonify({
            "ok": True,
            "type": kind,
            "message": response_text,
            "changed": changed,
            "stats": _doc_stats(core.current_doc),
            "preview_url": _preview_url(),
            "can_undo": bool(core.UNDO_STACK),
            "can_redo": bool(core.REDO_STACK),
        })


@app.post("/api/action")
def action():
    payload = request.get_json(silent=True) or {}
    action_name = str(payload.get("action", "")).lower()
    with EDITOR_LOCK:
        state = _get_state()
        _sync_active(state)
        if core.current_doc is None:
            return jsonify({"error": "Upload a DOCX document first."}), 400
        before = copy.deepcopy(core.current_doc)
        if action_name == "undo":
            _, output = _capture(core.undo_document)
        elif action_name == "redo":
            _, output = _capture(core.redo_document)
        elif action_name == "reset":
            _, output = _capture(core.reset_document)
        elif action_name == "summarize":
            _, output = _capture(core.summarize_document)
        elif action_name == "inspect":
            _, output = _capture(core.inspect_document)
        elif action_name == "improve":
            _, output = _capture(core.improve_format)
        else:
            return jsonify({"error": "Unknown action."}), 400
        _store_active(state)
        return jsonify({
            "ok": True,
            "message": output or "Action completed.",
            "stats": _doc_stats(core.current_doc),
            "preview": _preview(core.current_doc),
            "preview_url": _preview_url(),
            "can_undo": bool(core.UNDO_STACK),
            "can_redo": bool(core.REDO_STACK),
        })


@app.get("/api/session-file/<path:filename>")
def session_file(filename):
    safe = os.path.basename(filename)
    root = _session_dir() / "uploads"
    path = root / safe
    if not path.exists() or not path.is_file():
        return jsonify({"error": "File not found."}), 404
    return send_file(path, as_attachment=True, download_name=safe)


@app.get("/api/export")
def export_documents():
    """Export the active document, or every loaded DOCX as a ZIP bundle."""
    import zipfile as _zipfile
    with EDITOR_LOCK:
        state = _get_state()
        _sync_active(state)
        docs = list(state.get("documents", {}).items())
        if not docs:
            return jsonify({"error": "No DOCX documents are loaded."}), 404

        export_dir = _session_dir() / "exports"
        export_dir.mkdir(parents=True, exist_ok=True)

        # Save every in-memory document before exporting.
        for key, doc_state in docs:
            doc = doc_state.get("doc")
            path = doc_state.get("path")
            if doc is not None and path:
                doc.save(path)
        active = _active_state(state)
        if len(docs) == 1 and active:
            path = active.get("path")
            return send_file(path, as_attachment=True, download_name=os.path.basename(path))

        bundle = export_dir / f"edited-documents-{time.strftime('%Y%m%d-%H%M%S')}.zip"
        used = set()
        with _zipfile.ZipFile(bundle, "w", _zipfile.ZIP_DEFLATED) as zf:
            for key, doc_state in docs:
                path = doc_state.get("path")
                if not path or not os.path.isfile(path):
                    continue
                name = os.path.basename(path)
                if name in used:
                    name = f"{key}-{name}"
                used.add(name)
                zf.write(path, arcname=name)
        return send_file(bundle, as_attachment=True, download_name=bundle.name)


@app.get("/download/<path:filename>")
def download(filename):
    safe = os.path.basename(filename)
    path = OUTPUT_DIR / safe
    if not path.exists():
        return jsonify({"error": "Download expired or file not found."}), 404
    return send_file(path, as_attachment=True, download_name=safe.split("_", 1)[-1])


if __name__ == "__main__":
    host = os.environ.get("WEB_HOST", "127.0.0.1")
    port = int(os.environ.get("WEB_PORT", "5000"))
    app.run(host=host, port=port, debug=False)
