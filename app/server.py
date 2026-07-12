#!/usr/bin/env python3
"""
Artifact Generator — production service (POC v3 productionized).

Multi-format (pptx/docx/xlsx) generator. Returns a Dify-compatible artifact
contract (JSON metadata) so KPRX / Dify workflows consume it unchanged.

Endpoints:
    POST /generate/{fmt}/metadata   -> JSON artifact contract (Dify-compatible)
    POST /generate/{fmt}            -> binary (legacy, direct HTTP testing)
    GET  /download/{fmt}/{fname}    -> serve stored artifact
    GET  /health                    -> {"status":"ok",...}

Contract (unchanged from POC) — artifact metadata fields:
    provider / provider_file_id / download_url / filename / mime_type /
    artifact_type / size / checksum

Behavior is preserved from the POC; productionization changes are:
  * bind host configurable (default 172.17.0.1, NOT 0.0.0.0) — Dify-reachable
  * port 9191 (matches production)
  * threaded server for concurrency
  * output / template dirs configurable via env
"""
import argparse
import base64
import hashlib
import io
import json
import os
import sys
import tempfile
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote

# ---------------------------------------------------------------------------
# Configuration (env-overridable; defaults match production scope)
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.environ.get("AG_OUTPUT_DIR", os.path.join(BASE_DIR, "output"))
TEMPLATES_DIR = os.environ.get("AG_TEMPLATES_DIR", os.path.join(BASE_DIR, "templates"))
BIND_HOST = os.environ.get("AG_BIND_HOST", "172.17.0.1")
# External base URL used in download_url of the contract. In production the
# Dify container reaches the host via 172.17.0.1:9191; overridable for tests.
PUBLIC_BASE = os.environ.get("AG_PUBLIC_BASE", "http://172.17.0.1:9191")
DEFAULT_PORT = 9191

TEMPLATES = {
    "pptx": os.path.join(TEMPLATES_DIR, "synthetic_template.pptx"),
    "docx": os.path.join(TEMPLATES_DIR, "synthetic_template.docx"),
    "xlsx": os.path.join(TEMPLATES_DIR, "synthetic_template.xlsx"),
}
# default template_mode may only read from these roots (path-traversal guard)
ALLOWED_TEMPLATE_DIRS = [TEMPLATES_DIR, tempfile.gettempdir()]
os.makedirs(OUTPUT_DIR, exist_ok=True)

MIME_MAP = {
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}


# ---------------------------------------------------------------------------
# Template Mode Resolver (behavior preserved from POC)
# ---------------------------------------------------------------------------
def resolve_template(body: dict, fmt: str):
    mode = body.get("template_mode", "default")
    if mode == "user_upload":
        b64 = body.get("template_base64", "")
        if not b64:
            raise ValueError("user_upload requires template_base64")
        raw = base64.b64decode(b64)
        suffix = ".pptx" if fmt == "pptx" else ".docx" if fmt == "docx" else ".xlsx"
        tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        tmp.write(raw)
        tmp.close()
        return tmp.name, "[User Upload Template]"
    elif mode == "default":
        tp = body.get("template_path") or TEMPLATES.get(fmt)
        if not tp or not os.path.exists(tp):
            raise ValueError(f"default template not found: {tp}")
        if not any(
            os.path.realpath(tp).startswith(os.path.realpath(d) + os.sep)
            or os.path.realpath(tp) == os.path.realpath(d)
            for d in ALLOWED_TEMPLATE_DIRS
        ):
            raise ValueError(f"template path outside allowed dirs: {tp}")
        return tp, "[Default Template]"
    elif mode == "no_template":
        return None, "[No Template]"
    else:
        raise ValueError(f"unknown template_mode: {mode}")


# ---------------------------------------------------------------------------
# Generators (logic preserved from POC)
# ---------------------------------------------------------------------------
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


def generate_pptx(content, template_path, marker):
    prs = Presentation(template_path) if template_path else Presentation()
    sections = content.get(
        "sections", [{"heading": content.get("title", "Untitled"), "body": "", "bullets": []}]
    )
    for i, sec in enumerate(sections):
        if i == 0:
            if template_path:
                sl = prs.slides.add_slide(
                    prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
                )
                try:
                    sl.shapes.title.text = marker + " — " + sec["heading"]
                except Exception:
                    pass
            else:
                sl = prs.slides.add_slide(prs.slide_layouts[6])
                tb = sl.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11), Inches(3))
                tf = tb.text_frame
                p = tf.add_paragraph()
                p.text = marker
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                p.alignment = PP_ALIGN.CENTER
                p = tf.add_paragraph()
                p.text = sec["heading"]
                p.font.size = Pt(36)
                p.alignment = PP_ALIGN.CENTER
        else:
            sl = prs.slides.add_slide(prs.slide_layouts[1])
            try:
                sl.shapes.title.text = sec["heading"]
            except Exception:
                pass
            try:
                bf = sl.placeholders[1].text_frame
                if sec.get("body"):
                    p = bf.add_paragraph()
                    p.text = sec["body"]
                    p.font.size = Pt(18)
                for b in sec.get("bullets", []):
                    p = bf.add_paragraph()
                    p.text = "• " + b
                    p.font.size = Pt(16)
                    p.space_after = Pt(8)
            except Exception:
                pass
    for sl in prs.slides:
        try:
            tb = sl.shapes.add_textbox(Inches(0.5), Inches(7.0), Inches(12), Inches(0.4))
            tf = tb.text_frame
            p = tf.add_paragraph()
            p.text = marker
            p.font.size = Pt(8)
            p.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        except Exception:
            pass
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


from docx import Document
from docx.shared import Pt as DPt  # noqa: F401  (preserved import for parity)
from docx.enum.text import WD_ALIGN_PARAGRAPH


def generate_docx(content, template_path, marker):
    doc = Document(template_path) if template_path else Document()
    marker_para = doc.add_paragraph(marker)
    marker_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sections = content.get(
        "sections", [{"heading": content.get("title", "Untitled"), "body": "", "bullets": []}]
    )
    for sec in sections:
        doc.add_heading(sec["heading"], level=1)
        if sec.get("body"):
            doc.add_paragraph(sec["body"])
        for b in sec.get("bullets", []):
            doc.add_paragraph(b, style="List Bullet")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


def generate_xlsx(content, template_path, marker):
    wb = load_workbook(template_path) if template_path else Workbook()
    ws = wb.active
    ws.title = (content.get("title", "Data")[:28] + "-" + content.get("template_mode", "default")[:2])[:31]
    ws.insert_rows(1)
    ws.cell(1, 1, marker).font = Font(size=10, color="999999")
    if content.get("table_data"):
        td = content["table_data"]
        start = ws.max_row + 1
        hf = Font(bold=True, size=12, color="FFFFFF")
        hf2 = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
        for i, h in enumerate(td["headers"], 1):
            c = ws.cell(row=start, column=i, value=h)
            c.font = hf
            c.fill = hf2
        for ri, row in enumerate(td["rows"]):
            for ci, v in enumerate(row):
                ws.cell(row=start + ri + 1, column=ci + 1, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


GENERATORS = {"pptx": generate_pptx, "docx": generate_docx, "xlsx": generate_xlsx}

# OOXML formats are ZIP containers -> magic bytes "PK\x03\x04"
MAGIC = b"PK\x03\x04"


# ---------------------------------------------------------------------------
# HTTP Server
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "ArtifactGenerator/1.0"

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    # ---- helpers ----
    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw)

    # ---- GET ----
    def do_GET(self):
        path = unquote(self.path)
        if path == "/health":
            return self._json(
                {"status": "ok", "service": "artifact-generator", "version": "1.0"}
            )
        parts = [p for p in path.split("/") if p != ""]
        # /download/{fmt}/{filename}
        if len(parts) >= 3 and parts[0] == "download":
            fmt, fname = parts[1], "/".join(parts[2:])
            fpath = os.path.join(OUTPUT_DIR, fname)
            # path-traversal guard: resolved path must stay under OUTPUT_DIR
            if not (
                os.path.realpath(fpath) == os.path.realpath(OUTPUT_DIR)
                or os.path.realpath(fpath).startswith(os.path.realpath(OUTPUT_DIR) + os.sep)
            ):
                return self._json({"error": "path traversal blocked"}, 403)
            if not os.path.exists(fpath):
                return self._json({"error": "file not found"}, 404)
            with open(fpath, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", MIME_MAP.get(fmt, "application/octet-stream"))
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        return self._json({"error": "not found"}, 404)

    # ---- POST ----
    def do_POST(self):
        path = unquote(self.path)
        parts = [p for p in path.split("/") if p != ""]
        # /generate/{fmt}[/metadata]
        if len(parts) < 2 or parts[0] != "generate":
            return self._json({"error": "not found"}, 404)
        is_metadata = len(parts) >= 3 and parts[-1] == "metadata"
        fmt = parts[-2] if is_metadata else parts[-1]

        if fmt not in GENERATORS:
            return self._json({"error": f"unknown format: {fmt}"}, 400)

        try:
            body = self._read_json_body()
        except Exception:
            return self._json({"error": "invalid JSON body"}, 400)

        try:
            tmpl_path, marker = resolve_template(body, fmt)
            raw = GENERATORS[fmt](body, tmpl_path, marker)
        except ValueError as e:
            return self._json({"error": str(e)}, 400)
        except Exception as e:
            return self._json({"error": f"generation failed: {e}"}, 500)

        if is_metadata:
            sha = hashlib.sha256(raw).hexdigest()
            fname = f"sandbox-{fmt}-{sha[:12]}.{fmt}"
            fpath = os.path.join(OUTPUT_DIR, fname)
            with open(fpath, "wb") as f:
                f.write(raw)
            title = body.get("title", "Untitled")
            contract = {
                "artifact": {
                    "provider": "external_tool",
                    "provider_file_id": sha[:12],
                    "download_url": f"{PUBLIC_BASE}/download/{fmt}/{fname}",
                    "filename": f"{title}.{fmt}",
                    "mime_type": MIME_MAP[fmt],
                    "artifact_type": fmt,
                    "size": len(raw),
                    "checksum": f"sha256:{sha}",
                }
            }
            return self._json(contract)

        # binary (legacy)
        self.send_response(200)
        self.send_header("Content-Type", MIME_MAP[fmt])
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)


def main():
    ap = argparse.ArgumentParser(description="Artifact Generator production service")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument(
        "--host",
        default=BIND_HOST,
        help="bind host (default 172.17.0.1; Dify-reachable, not 0.0.0.0)",
    )
    args = ap.parse_args()
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Artifact Generator on {args.host}:{args.port} (output={OUTPUT_DIR})", flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.shutdown()


if __name__ == "__main__":
    main()
