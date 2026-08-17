"""
CRQ — thin HTTP layer for Step 0 + Layer 1 (stdlib only).

Serves the static frontend (static/index.html) and two JSON endpoints on
top of crq_api.py. Contains NO calculation logic of its own — all numbers
come from crq_api.calculate() (unchanged Step 1 contract) plus the
read-only get_top_threat() enrichment.

  GET  /api/options     -> dropdown data for Step 0 (Industry_Master display
                          names excl. orphans, Frequency_IRIS Table C labels
                          in workbook order, Frequency_SSC grades)
  POST /api/calculate   -> {industry, revenue, ssc_grade?} ->
                          calculate() result + "top_threat" enrichment;
                          Step 0 validation errors return HTTP 400 JSON.
  GET  /api/methodology -> Layer 4 data: Assumption Register, Evidence
                          Register, Industry_Master confidence tally
                          (static across all calculations, no inputs).
  POST /api/report/pdf  -> {html, filename?} -> renders the client-built
                          standalone report HTML through headless Chrome's
                          own --print-to-pdf flag and streams back the PDF
                          bytes as a download (no PDF library dependency;
                          no calculation logic — html is fully pre-rendered
                          client-side from an already-computed result).

Run:  python3 server.py   (then open http://localhost:8765)
"""

import json
import os
import subprocess
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from crq_api import ValidationError, _load_data, calculate, get_methodology, get_top_threat

PORT = int(os.environ.get("PORT", 8765))
_STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

# PDF report generation shells out to Chrome's own --print-to-pdf headless
# flag (2026-08-07, user-directed: "크롬 기준으로") instead of adding a PDF
# library dependency — consistent with this server staying stdlib-only.
_CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
]


def _find_chrome():
    for path in _CHROME_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


class Handler(BaseHTTPRequestHandler):

    def end_headers(self):
        # Frontend and backend deploy to different origins (Netlify + Render),
        # so every response needs CORS headers for the browser to accept it.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, content_type):
        with open(path, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in ("/", "/index.html"):
            self._send_file(os.path.join(_STATIC_DIR, "index.html"),
                            "text/html; charset=utf-8")
        elif path == "/api/options":
            data = _load_data()
            industries = [
                {
                    "industry_id": r["industry_id"],
                    "display_name": r["display_name"],
                    "mapping_confidence": r["mapping_confidence"],
                }
                for r in data["industry_master"]
                if r["iris_frequency_name"] not in (None, "N/A")
            ]
            self._send_json({
                "industries": industries,
                # dict order == workbook row order (Table C top-to-bottom)
                "revenue_classes": list(data["rev_mult"]),
                "ssc_grades": list(data["ssc_mult"]),
            })
        elif path == "/api/methodology":
            self._send_json(get_methodology())
        else:
            self._send_json({"error": "not found"}, status=404)

    def do_POST(self):
        path = self.path.split("?", 1)[0]
        if path == "/api/calculate":
            self._handle_calculate()
        elif path == "/api/report/pdf":
            self._handle_report_pdf()
        else:
            self._send_json({"error": "not found"}, status=404)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length) or b"{}")

    def _handle_calculate(self):
        try:
            payload = self._read_json_body()
        except (ValueError, json.JSONDecodeError):
            self._send_json({"error": "invalid JSON body"}, status=400)
            return
        try:
            result = calculate(
                payload.get("industry"),
                payload.get("revenue"),
                ssc_grade=payload.get("ssc_grade") or None,
            )
            result["top_threat"] = get_top_threat(payload.get("industry"))
        except ValidationError as e:
            self._send_json({"error": str(e)}, status=400)
            return
        self._send_json(result)

    def _handle_report_pdf(self):
        try:
            payload = self._read_json_body()
        except (ValueError, json.JSONDecodeError):
            self._send_json({"error": "invalid JSON body"}, status=400)
            return
        html = payload.get("html")
        filename = payload.get("filename") or "report.pdf"
        if not html:
            self._send_json({"error": "html is required"}, status=400)
            return
        chrome = _find_chrome()
        if not chrome:
            self._send_json({"error": "Chrome executable not found on this machine"},
                             status=500)
            return
        with tempfile.TemporaryDirectory() as tmpdir:
            html_path = os.path.join(tmpdir, "report.html")
            pdf_path = os.path.join(tmpdir, "report.pdf")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
            try:
                subprocess.run(
                    [chrome, "--headless=new", "--disable-gpu", "--no-sandbox",
                     "--print-to-pdf=" + pdf_path, "--no-pdf-header-footer",
                     "file://" + html_path],
                    check=True, timeout=20,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
                self._send_json({"error": "PDF generation failed"}, status=500)
                return
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
        self.send_response(200)
        self.send_header("Content-Type", "application/pdf")
        self.send_header("Content-Disposition",
                          'attachment; filename="' + filename + '"')
        self.send_header("Content-Length", str(len(pdf_bytes)))
        self.end_headers()
        self.wfile.write(pdf_bytes)

    def log_message(self, fmt, *args):  # quiet default access log
        pass


if __name__ == "__main__":
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"CRQ serving on http://localhost:{PORT}")
    server.serve_forever()
