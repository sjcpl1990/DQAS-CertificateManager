"""
DQAS Certificate Manager
------------------------
A Streamlit app to process DQAS-qualified personnel certificates
(sourced as individual PDFs or combined trade-wise multi-page PDFs),
automatically extract each person's details, rename them to our
certificate-ID format, generate a QR code that checks validity on every
scan, and stamp the QR onto the certificate automatically. Organized by
trade (global — no per-site grouping).
"""

import os
import re
import io
import json
import shutil
import secrets
import zipfile
from datetime import datetime, date
from html import escape as html_escape

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

import common

# The app is self-contained: it creates its own database and a default
# certificates folder automatically, no external paths to configure to get
# started. The certificates folder defaults to right next to this script —
# ordinary file read/write there is fine wherever the repo is checked out.
#
# The *database* deliberately does NOT default next to the script: SQLite
# opens and closes a new connection per operation, which relies on real
# file-locking semantics that cloud-sync/network drives (OneDrive, Google
# Drive, a mapped network share — confirmed on this machine's Z:\) can
# silently break, corrupting the file ("database disk image is malformed")
# after just a few writes. A per-user local-appdata folder is guaranteed to
# be on a real local disk regardless of where this repo itself lives.
_APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _local_data_dir():
    base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
    data_dir = os.path.join(base, "DQASCertificateManager")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir


DB_PATH = os.path.join(_local_data_dir(), "dqas_certificates.db")
DEFAULT_CERTS_ROOT = os.path.join(_APP_DIR, "DQAS Certificates")

APP_TITLE = "DQAS Certificate Manager"
APP_CAPTION = "DQAS Certificate Manager — Birla Punya"

EXCEL_COLUMNS = [
    ("Certificate ID", "cert_id"), ("Person Name", "person_name"), ("Employee ID", "employee_id"),
    ("Trade", "trade_code"), ("Designation", "designation"), ("Issue Date", "issue_date"),
    ("Expiry Date", "expiry_date"),
]

MONTH_MAP = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


# ----------------------------------------------------------------------
# Database layer
# ----------------------------------------------------------------------

def init_db():
    conn = common.get_conn(DB_PATH)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS dqas_certificates (
            cert_id TEXT PRIMARY KEY,
            person_name TEXT NOT NULL,
            employee_id TEXT,
            trade_code TEXT NOT NULL,
            designation TEXT,
            certificate_no TEXT,
            link TEXT,
            source_path TEXT,
            renamed_path TEXT,
            stamped_path TEXT,
            issue_date TEXT NOT NULL,
            expiry_date TEXT NOT NULL,
            status TEXT DEFAULT 'Active',
            verify_token TEXT,
            created_at TEXT,
            updated_at TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    conn.execute("CREATE TABLE IF NOT EXISTS trades (code TEXT PRIMARY KEY, name TEXT)")

    # Migrate: static-mode URLs use a random token, not the (guessable,
    # sequential) certificate ID, so pages can't be enumerated by anyone
    # who finds one. Add the column and backfill any rows that predate it.
    existing_cols = [r[1] for r in conn.execute("PRAGMA table_info(dqas_certificates)").fetchall()]
    if "verify_token" not in existing_cols:
        conn.execute("ALTER TABLE dqas_certificates ADD COLUMN verify_token TEXT")
    for (cid,) in conn.execute("SELECT cert_id FROM dqas_certificates WHERE verify_token IS NULL OR verify_token = ''").fetchall():
        conn.execute("UPDATE dqas_certificates SET verify_token=? WHERE cert_id=?", (secrets.token_urlsafe(12), cid))

    conn.commit()
    conn.close()


# --- Trades lookup (code -> friendly display name) ---

def get_trades():
    conn = common.get_conn(DB_PATH)
    rows = conn.execute("SELECT code, name FROM trades ORDER BY name").fetchall()
    conn.close()
    return rows


def get_trade_name(code):
    conn = common.get_conn(DB_PATH)
    row = conn.execute("SELECT name FROM trades WHERE code=?", (code,)).fetchone()
    conn.close()
    return row[0] if row else code


def upsert_trade(code, name):
    conn = common.get_conn(DB_PATH)
    conn.execute(
        "INSERT INTO trades (code, name) VALUES (?, ?) "
        "ON CONFLICT(code) DO UPDATE SET name=excluded.name",
        (code, name),
    )
    conn.commit()
    conn.close()


def delete_trade(code):
    conn = common.get_conn(DB_PATH)
    conn.execute("DELETE FROM trades WHERE code=?", (code,))
    conn.commit()
    conn.close()


# --- DQAS Certificate CRUD ---

def next_cert_id(trade_code):
    prefix = f"CERT-DQAS-{trade_code}-"
    conn = common.get_conn(DB_PATH)
    rows = conn.execute("SELECT cert_id FROM dqas_certificates WHERE cert_id LIKE ?", (f"{prefix}%",)).fetchall()
    conn.close()
    max_seq = 0
    for (cid,) in rows:
        suffix = cid[len(prefix):]
        if suffix.isdigit():
            max_seq = max(max_seq, int(suffix))
    return f"{prefix}{max_seq + 1:02d}"


def add_certificate(cert_id, person_name, employee_id, trade_code, designation, certificate_no,
                     link, source_path, renamed_path, stamped_path, issue_date, expiry_date):
    now = datetime.now().isoformat(timespec="seconds")
    token = secrets.token_urlsafe(12)
    conn = common.get_conn(DB_PATH)
    conn.execute(
        "INSERT INTO dqas_certificates (cert_id, person_name, employee_id, trade_code, designation, "
        "certificate_no, link, source_path, renamed_path, stamped_path, issue_date, expiry_date, "
        "status, verify_token, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active', ?, ?, ?)",
        (cert_id, person_name, employee_id, trade_code, designation, certificate_no, link,
         source_path, renamed_path, stamped_path, issue_date, expiry_date, token, now, now),
    )
    conn.commit()
    conn.close()


def rotate_verify_token(cert_id):
    """Issues a fresh random token for a certificate's static verification
    URL, invalidating the old one — e.g. if a printed QR is compromised."""
    token = secrets.token_urlsafe(12)
    conn = common.get_conn(DB_PATH)
    conn.execute("UPDATE dqas_certificates SET verify_token=? WHERE cert_id=?", (token, cert_id))
    conn.commit()
    conn.close()
    return token


def update_certificate(cert_id, person_name, employee_id, trade_code, designation, certificate_no,
                        link, source_path, renamed_path, stamped_path, issue_date, expiry_date, status):
    now = datetime.now().isoformat(timespec="seconds")
    conn = common.get_conn(DB_PATH)
    conn.execute(
        "UPDATE dqas_certificates SET person_name=?, employee_id=?, trade_code=?, designation=?, "
        "certificate_no=?, link=?, source_path=?, renamed_path=?, stamped_path=?, issue_date=?, "
        "expiry_date=?, status=?, updated_at=? WHERE cert_id=?",
        (person_name, employee_id, trade_code, designation, certificate_no, link, source_path,
         renamed_path, stamped_path, issue_date, expiry_date, status, now, cert_id),
    )
    conn.commit()
    conn.close()


def delete_certificate(cert_id):
    conn = common.get_conn(DB_PATH)
    conn.execute("DELETE FROM dqas_certificates WHERE cert_id=?", (cert_id,))
    conn.commit()
    conn.close()


def get_certificates(trade_filter=None, status_filter=None, search=None):
    conn = common.get_conn(DB_PATH)
    query = "SELECT * FROM dqas_certificates WHERE 1=1"
    params = []
    if trade_filter and trade_filter != "All":
        query += " AND trade_code=?"
        params.append(trade_filter)
    if status_filter and status_filter != "All":
        query += " AND status=?"
        params.append(status_filter)
    if search:
        query += " AND (cert_id LIKE ? OR person_name LIKE ? OR employee_id LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
    query += " ORDER BY trade_code, person_name"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_certificate(cert_id):
    conn = common.get_conn(DB_PATH)
    row = conn.execute("SELECT * FROM dqas_certificates WHERE cert_id=?", (cert_id,)).fetchone()
    cols = [d[0] for d in conn.execute("SELECT * FROM dqas_certificates LIMIT 0").description]
    conn.close()
    if not row:
        return None
    return dict(zip(cols, row))


def find_existing_certificate(certificate_no, employee_id, trade_code, issue_date):
    conn = common.get_conn(DB_PATH)
    row = None
    if certificate_no:
        row = conn.execute(
            "SELECT cert_id FROM dqas_certificates WHERE certificate_no=?", (certificate_no,)
        ).fetchone()
    if not row and employee_id:
        row = conn.execute(
            "SELECT cert_id FROM dqas_certificates WHERE employee_id=? AND trade_code=? AND issue_date=?",
            (employee_id, trade_code, issue_date),
        ).fetchone()
    conn.close()
    return row[0] if row else None


# ----------------------------------------------------------------------
# QR target — what the QR code actually encodes, per the configured
# link mode:
#   "static"   — a link to a self-contained static HTML page (generated by
#                build_static_verification_html) that checks validity in the
#                visitor's own browser. No app hosting required; the page
#                must be hosted somewhere static (e.g. GitHub Pages).
#   "direct"   — the certificate's own link, straight through. No validity
#                check at scan time — simplest, needs nothing hosted at all.
#   "redirect" — a link back to this app (?cert=<id>), which checks validity
#                server-side. Requires this app to be deployed at a
#                permanent public URL.
# ----------------------------------------------------------------------

def build_cert_qr_target(cert_id, cert=None):
    # cert may be a dict, a pandas Series (from a DataFrame row), or None —
    # never test it for truthiness directly (a Series' truth value is
    # ambiguous); check "is None" and use .get(), which both dict and
    # Series support identically.
    mode = common.get_setting(DB_PATH, "link_mode", "static")

    if mode == "direct":
        if cert is None:
            cert = get_certificate(cert_id)
        return (cert.get("link") or None) if cert is not None else None

    if mode == "static":
        pages_base = common.get_setting(DB_PATH, "pages_base_url", "").rstrip("/")
        if not pages_base:
            return None
        if cert is None:
            cert = get_certificate(cert_id)
        token = cert.get("verify_token") if cert is not None else None
        if not token:
            return None
        # A random per-certificate token, not the (sequential, guessable)
        # cert_id — so finding one page's URL doesn't let anyone enumerate
        # every other certificate.
        return f"{pages_base}/{token}.html"

    base_url = common.get_setting(DB_PATH, "base_url", "").rstrip("/")
    if not base_url:
        return None
    return f"{base_url}/?cert={cert_id}"


def qr_config_status():
    """(ready, warning_or_None) for the currently configured link mode —
    used to decide whether to show a 'can't generate QR yet' warning."""
    mode = common.get_setting(DB_PATH, "link_mode", "static")
    if mode == "direct":
        return True, None
    if mode == "static":
        if common.get_setting(DB_PATH, "pages_base_url", "").strip():
            return True, None
        return False, (
            "⚠️ Static pages base URL is not set. Set it in the Settings tab (Link mode → "
            "Static verification page) before generating QR codes."
        )
    if common.get_setting(DB_PATH, "base_url", "").strip():
        return True, None
    return False, (
        "⚠️ App base URL is not set. Set it in the Settings tab (Link mode → App redirect) "
        "before generating QR codes."
    )


def apply_signature_if_set(img):
    """Overlays the GM signature (if one has been uploaded and locked) onto
    img at its saved position. Returns img unchanged if no signature is set."""
    sig_path = common.get_setting(DB_PATH, "signature_path", "")
    if not sig_path or not os.path.exists(sig_path):
        return img
    sig_img = common.load_overlay_image(sig_path)
    sx = float(common.get_setting(DB_PATH, "sig_x_pct", "50"))
    sy = float(common.get_setting(DB_PATH, "sig_y_pct", "60"))
    ssize = float(common.get_setting(DB_PATH, "sig_size_pct", "15"))
    return common.paste_overlay(img, sig_img, sx, sy, ssize)


# ----------------------------------------------------------------------
# Verification rendering — shared by the public scan handler and the
# in-app Test tab (so "test the scan" shows exactly what a phone would see)
# ----------------------------------------------------------------------

def render_verification_html(cert, as_of_date, redirect=True):
    trade_label = cert.get("designation") or cert.get("trade_code") or "—"

    if cert["status"] == "Revoked":
        return f"""
        <div style="background:#ffe5e5;border:2px solid #d32f2f;border-radius:8px;padding:24px;text-align:center;">
        <h2 style="color:#d32f2f;">❌ Certificate Revoked</h2>
        <p style="font-size:18px;"><strong>{cert['person_name']}</strong> — {trade_label}<br>
        Employee ID: {cert.get('employee_id') or '—'}<br>
        Certificate ID: {cert['cert_id']}</p>
        <p>This certificate has been revoked and is no longer valid.</p>
        </div>
        """

    expiry = datetime.strptime(cert["expiry_date"], "%Y-%m-%d").date()

    if as_of_date > expiry:
        return f"""
        <div style="background:#ffe5e5;border:2px solid #d32f2f;border-radius:8px;padding:24px;text-align:center;">
        <h2 style="color:#d32f2f;">❌ Certificate Validity Expired</h2>
        <p style="font-size:18px;"><strong>{cert['person_name']}</strong> — {trade_label}<br>
        Employee ID: {cert.get('employee_id') or '—'}<br>
        Certificate ID: {cert['cert_id']}</p>
        <p>This certificate was valid until <strong>{expiry.isoformat()}</strong> and has now expired.</p>
        <p>Please contact the DQAS administrator for a renewed certificate.</p>
        </div>
        """

    link = cert.get("link") or ""
    meta = f'<meta http-equiv="refresh" content="2; url={link}">' if redirect and link else ""
    if redirect and link:
        note = "Redirecting to the certificate…"
    elif link:
        note = ""
    else:
        note = "No certificate link has been set yet."
    return f"""
    <div style="background:#e6f7ec;border:2px solid #2e7d32;border-radius:8px;padding:24px;text-align:center;">
    <h2 style="color:#2e7d32;">✅ Valid Certificate</h2>
    <p style="font-size:18px;"><strong>{cert['person_name']}</strong> — {trade_label}<br>
    Employee ID: {cert.get('employee_id') or '—'}<br>
    Certificate ID: {cert['cert_id']}</p>
    <p>Valid until <strong>{expiry.isoformat()}</strong>. {note}</p>
    </div>
    {meta}
    """


def build_static_verification_html(cert):
    """A self-contained HTML+JS page for 'static' link mode: it checks
    today's date against the expiry date in the *visitor's own browser* —
    no server or hosted app involved — then shows a Valid/Expired banner
    and, if valid, a button to the certificate.

    Limitation this can't get around: revocation is a manual, arbitrary
    action, not something a date comparison can detect. A certificate
    revoked after this page was generated won't show as revoked until the
    page is regenerated and reshared — the page says so explicitly."""
    trade_label = cert.get("designation") or cert.get("trade_code") or "—"
    link = cert.get("link") or ""
    is_revoked = cert["status"] == "Revoked"
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    rows = [
        ("Name", cert["person_name"]),
        ("Trade", trade_label),
        ("Employee ID", cert.get("employee_id") or "—"),
        ("Certificate ID", cert["cert_id"]),
        ("Issued", cert["issue_date"]),
        ("Valid Until", cert["expiry_date"]),
    ]
    meta_html = "\n".join(
        f'<div><b>{html_escape(k)}:</b> {html_escape(str(v))}</div>' for k, v in rows
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Certificate Verification — {html_escape(cert['cert_id'])}</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif; background:#f4f6f8; margin:0; padding:24px; }}
  .card {{ max-width:480px; margin:40px auto; background:#fff; border-radius:12px; box-shadow:0 2px 12px rgba(0,0,0,.08); padding:32px; text-align:center; }}
  .banner {{ border-radius:8px; padding:20px; margin-bottom:20px; }}
  .valid {{ background:#e6f7ec; border:2px solid #2e7d32; }}
  .bad {{ background:#ffe5e5; border:2px solid #d32f2f; }}
  h2 {{ margin:0 0 8px; font-size:22px; }}
  .valid h2 {{ color:#2e7d32; }}
  .bad h2 {{ color:#d32f2f; }}
  .meta {{ text-align:left; font-size:14px; color:#333; margin-top:16px; line-height:1.7; }}
  .meta b {{ color:#000; }}
  a.btn {{ display:inline-block; margin-top:18px; padding:10px 26px; background:#00AEDA; color:#fff; text-decoration:none; border-radius:6px; font-weight:600; }}
  .note {{ font-size:12px; color:#888; margin-top:22px; line-height:1.5; }}
</style>
</head>
<body>
<div class="card">
  <div id="banner" class="banner"><h2 id="status-title">Checking…</h2><p id="status-text"></p></div>
  <div class="meta">{meta_html}</div>
  <div id="link-area"></div>
  <p class="note">
    Generated {html_escape(generated_at)}. This page checks the date in your own browser — no server involved.<br>
    If this certificate is later revoked, that change won't appear here until this page is regenerated and reshared.
  </p>
</div>
<script>
(function () {{
  var expiry = new Date({json.dumps(cert['expiry_date'] + "T23:59:59")});
  var certLink = {json.dumps(link)};
  var isRevoked = {json.dumps(is_revoked)};
  var validUntil = {json.dumps(cert['expiry_date'])};
  var banner = document.getElementById("banner");
  var title = document.getElementById("status-title");
  var text = document.getElementById("status-text");
  var linkArea = document.getElementById("link-area");

  if (isRevoked) {{
    banner.className = "banner bad";
    title.textContent = "\\u274c Certificate Revoked";
    text.textContent = "This certificate had been revoked as of when this page was last generated.";
  }} else if (new Date() > expiry) {{
    banner.className = "banner bad";
    title.textContent = "\\u274c Certificate Validity Expired";
    text.textContent = "This certificate was valid until " + validUntil + " and has now expired.";
  }} else {{
    banner.className = "banner valid";
    title.textContent = "\\u2705 Valid Certificate";
    text.textContent = "Valid until " + validUntil + ".";
    if (certLink) {{
      linkArea.innerHTML = '<a class="btn" href="' + certLink + '">View Certificate</a>';
    }} else {{
      text.textContent += " (Certificate link not set yet.)";
    }}
  }}
}})();
</script>
</body>
</html>
"""


def handle_cert_redirect():
    cert = get_certificate(st.query_params.get("cert"))
    st.set_page_config(page_title="Certificate Verification", page_icon="🪪")
    if not cert:
        st.error("No certificate found for that ID.")
        st.stop()

    today = datetime.now().date()
    html = render_verification_html(cert, today, redirect=True)

    extra_link = ""
    if cert["status"] != "Revoked":
        expiry = datetime.strptime(cert["expiry_date"], "%Y-%m-%d").date()
        if today <= expiry and cert.get("link"):
            extra_link = f'<p>If not redirected, <a href="{cert["link"]}">click here</a>.</p>'

    st.markdown(html + extra_link, unsafe_allow_html=True)
    st.stop()


# ----------------------------------------------------------------------
# Certificate text extraction
# ----------------------------------------------------------------------

def parse_cert_date(s):
    """Parse dates like '01 Sept 2026'. Returns (date_or_None, error_or_None)."""
    s = (s or "").strip()
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if not m:
        return None, f"Unrecognized date format: '{s}'"
    day, mon_str, year = m.groups()
    mon = MONTH_MAP.get(mon_str.lower())
    if not mon:
        return None, f"Unrecognized month: '{mon_str}'"
    try:
        return date(int(year), mon, int(day)), None
    except ValueError as e:
        return None, f"Invalid date '{s}': {e}"


def extract_cert_fields(text):
    """Pulls Name / Employee ID / Designation / Certificate No. / dates out
    of a certificate page's extracted text, and derives a trade code."""
    name_m = re.search(r"This is to certify that\s*\n?\s*(.+?)\s*\n\s*has successfully completed", text, re.DOTALL)
    person_name = name_m.group(1).strip() if name_m else ""

    emp_m = re.search(r"Employee ID:\s*(\S+)", text)
    employee_id = emp_m.group(1).strip() if emp_m else ""

    des_m = re.search(r"Designation:\s*(.+)", text)
    designation = des_m.group(1).strip() if des_m else ""

    certno_m = re.search(r"Certificate No\.?:\s*(.+)", text)
    certificate_no = certno_m.group(1).strip() if certno_m else ""

    issue_m = re.search(r"Issue Date:\s*(.+)", text)
    issue_raw = issue_m.group(1).strip() if issue_m else ""

    valid_m = re.search(r"Valid Until:\s*(.+)", text)
    expiry_raw = valid_m.group(1).strip() if valid_m else ""

    code_m = re.match(r"DQAS-([A-Za-z0-9]+)-", certificate_no)
    if code_m:
        trade_code = code_m.group(1).upper()
    else:
        words = re.findall(r"[A-Za-z]+", designation)
        trade_code = ("".join(w[0] for w in words).upper()[:5]) or "GEN"

    issue_date, issue_err = parse_cert_date(issue_raw)
    expiry_date, expiry_err = parse_cert_date(expiry_raw)

    return {
        "person_name": person_name,
        "employee_id": employee_id,
        "designation": designation,
        "certificate_no": certificate_no,
        "trade_code": trade_code,
        "issue_date": issue_date.isoformat() if issue_date else "",
        "expiry_date": expiry_date.isoformat() if expiry_date else "",
        "date_error": issue_err or expiry_err or "",
    }


def scan_source(path):
    """path: a folder (scanned non-recursively, so subfolders already
    processed separately are never double-counted) or a single PDF file.
    Each PDF page found is treated as one certificate. Returns (rows, errors)."""
    if os.path.isdir(path):
        files = sorted(
            os.path.join(path, f) for f in os.listdir(path)
            if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(path, f))
        )
    elif os.path.isfile(path) and path.lower().endswith(".pdf"):
        files = [path]
    else:
        return [], ["Path is not a folder or a PDF file."]

    if not files:
        return [], ["No PDF files found at that location."]

    rows, errors = [], []
    for fpath in files:
        try:
            doc = fitz.open(fpath)
        except Exception as e:
            errors.append(f"{os.path.basename(fpath)}: could not open ({e})")
            continue
        for page_idx in range(doc.page_count):
            text = doc[page_idx].get_text()
            f = extract_cert_fields(text)
            if not f["person_name"]:
                errors.append(f"{os.path.basename(fpath)} (page {page_idx + 1}): couldn't find a name — skipped.")
                continue
            dup = find_existing_certificate(f["certificate_no"], f["employee_id"], f["trade_code"], f["issue_date"])
            if dup:
                status = f"Already imported ({dup})"
            elif f["date_error"]:
                status = f"⚠ Fix dates — {f['date_error']}"
            else:
                status = "New"
            rows.append({
                "Include": dup is None and not f["date_error"],
                "Person Name": f["person_name"],
                "Employee ID": f["employee_id"],
                "Trade Code": f["trade_code"],
                "Designation": f["designation"],
                "Issue Date": f["issue_date"],
                "Expiry Date": f["expiry_date"],
                "Status": status,
                "_file": fpath,
                "_page": page_idx,
                "_certificate_no": f["certificate_no"],
            })
        doc.close()
    return rows, errors


def safe_name(name):
    return "".join(c if c.isalnum() or c in " -_" else "_" for c in name).strip() or "Unknown"


def extract_single_page_pdf(fpath, page_idx, out_path):
    """Save one page of a (possibly multi-page) source PDF as its own file."""
    src = fitz.open(fpath)
    if src.page_count == 1 and page_idx == 0:
        src.close()
        shutil.copyfile(fpath, out_path)
        return
    new_doc = fitz.open()
    new_doc.insert_pdf(src, from_page=page_idx, to_page=page_idx)
    new_doc.save(out_path)
    new_doc.close()
    src.close()


# ----------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------

def main():
    init_db()

    # Public: QR scan verification — bypass login
    if "cert" in st.query_params:
        handle_cert_redirect()
        return

    # Everything else requires login
    if not st.session_state.get("authenticated"):
        common.show_login(APP_TITLE, APP_CAPTION, page_icon="🪪")
        return

    st.set_page_config(page_title=APP_TITLE, page_icon="🪪", layout="wide")

    top_l, top_r = st.columns([6, 1])
    with top_l:
        st.title("🪪 DQAS Certificate Manager")
        st.caption("Extract, rename, QR-stamp, and verify DQAS-qualified personnel certificates — global, trade-organized.")
    with top_r:
        st.write("")
        if st.button("Sign out"):
            st.session_state["authenticated"] = False
            st.rerun()

    tabs = st.tabs([
        "🪪 Certificates", "➕ Add Certificate", "⚙️🗂️ Process Certificates",
        "🖼️ Stamp Certificates", "🧪 Test", "🏷️ Manage Trades", "⚙️ Settings",
    ])

    # ── Certificates ───────────────────────────────────────────────────
    with tabs[0]:
        st.subheader("DQAS Certificates")
        st.caption(
            "QR codes always route through this app so validity can be checked on every "
            "scan — a certificate past its validity period shows an expired notice instead "
            "of opening the file."
        )

        qr_ready, qr_warning = qr_config_status()
        if qr_warning:
            st.warning(qr_warning)

        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            search = st.text_input("Search by ID, name, or employee ID", "")
        with col2:
            trade_codes = ["All"] + [c for c, _ in get_trades()]
            trade_filter = st.selectbox(
                "Trade", trade_codes,
                format_func=lambda c: c if c == "All" else f"{c} — {get_trade_name(c)}",
            )
        with col3:
            status_filter = st.selectbox("Status", ["All", "Active", "Revoked"])

        df = get_certificates(trade_filter, status_filter, search)

        if df.empty:
            st.info("No certificates yet. Add one in the 'Add Certificate' tab, or process a folder in 'Process Certificates'.")
        else:
            today = datetime.now().date()
            current_trade = None
            for _, row in df.iterrows():
                trade_label = f"{row['trade_code']} — {get_trade_name(row['trade_code'])}"
                if trade_label != current_trade:
                    current_trade = trade_label
                    st.markdown(f"### {trade_label}")

                expiry = datetime.strptime(row["expiry_date"], "%Y-%m-%d").date()
                is_expired = today > expiry
                is_revoked = row["status"] == "Revoked"
                data = build_cert_qr_target(row["cert_id"], cert=row)

                with st.container(border=True):
                    c1, c2, c3 = st.columns([1, 3, 1])
                    with c1:
                        if data:
                            st.image(common.make_qr_image(data, box_size=5), width=120)
                        else:
                            st.caption("Set base URL in Settings to generate QR")
                    with c2:
                        badge = "🔴 REVOKED" if is_revoked else ("🔴 EXPIRED" if is_expired else "🟢 VALID")
                        st.markdown(f"**{row['cert_id']} — {row['person_name']}** {badge}")
                        st.caption(
                            f"Employee ID: {row.get('employee_id') or '—'}  |  "
                            f"Designation: {row.get('designation') or '—'}  |  "
                            f"Issued: {row['issue_date']}  |  Expires: {row['expiry_date']}"
                        )
                        st.text(f"Link: {row['link'] or '(not set yet)'}")
                        if row.get("stamped_path"):
                            st.caption(f"📄 Stamped file: {row['stamped_path']}")
                        elif row.get("renamed_path"):
                            st.caption(f"📄 Renamed file: {row['renamed_path']}")
                    with c3:
                        if data:
                            labeled = common.make_labeled_qr(row["cert_id"], row["person_name"], data)
                            st.download_button(
                                "Download QR",
                                common.img_to_bytes(labeled),
                                file_name=f"{row['cert_id']}_qr.png",
                                mime="image/png",
                                key=f"dlc_{row['cert_id']}",
                            )
                        if common.get_setting(DB_PATH, "link_mode", "static") == "static":
                            st.download_button(
                                "Download verification page",
                                build_static_verification_html(row.to_dict()).encode("utf-8"),
                                file_name=f"{row.get('verify_token') or row['cert_id']}.html",
                                mime="text/html",
                                key=f"verifypage_{row['cert_id']}",
                                help="Re-download and re-upload this after editing the certificate (e.g. after adding its link), so the hosted page matches. Filename must match the URL the QR encodes.",
                            )
                            if st.button("🔄 Rotate verification link", key=f"rotate_{row['cert_id']}",
                                         help="Issues a new random URL for this certificate's static page and invalidates the old one — use if a printed QR/link was compromised. You'll need to reprint the QR and re-host the new .html file afterward."):
                                rotate_verify_token(row["cert_id"])
                                st.success("Link rotated — regenerate and re-stamp this certificate's QR, then re-upload its new verification page.")
                                st.rerun()
                        if common.zoho_configured(DB_PATH) and st.button("🔗 Get link from WorkDrive", key=f"zlink_{row['cert_id']}"):
                            file_for_link = row.get("stamped_path") or row.get("renamed_path")
                            if not file_for_link or not os.path.exists(file_for_link):
                                st.error("No local file available for this certificate to upload.")
                            else:
                                new_link, zerr = common.zoho_upload_and_link(DB_PATH, file_for_link)
                                if zerr:
                                    st.error(zerr)
                                else:
                                    update_certificate(
                                        row["cert_id"], row["person_name"], row["employee_id"], row["trade_code"],
                                        row["designation"], row["certificate_no"], new_link, row["source_path"],
                                        row["renamed_path"], row["stamped_path"], row["issue_date"], row["expiry_date"],
                                        row["status"],
                                    )
                                    st.success("Link fetched from WorkDrive.")
                                    st.rerun()
                        if not is_revoked and st.button("Revoke", key=f"revokec_{row['cert_id']}"):
                            update_certificate(
                                row["cert_id"], row["person_name"], row["employee_id"], row["trade_code"],
                                row["designation"], row["certificate_no"], row["link"], row["source_path"],
                                row["renamed_path"], row["stamped_path"], row["issue_date"], row["expiry_date"],
                                "Revoked",
                            )
                            st.rerun()
                        if st.button("Delete", key=f"delc_{row['cert_id']}"):
                            delete_certificate(row["cert_id"])
                            st.success(f"Deleted {row['cert_id']}")
                            st.rerun()

                    with st.expander("Edit"):
                        with st.form(f"edit_form_{row['cert_id']}"):
                            e1, e2 = st.columns(2)
                            with e1:
                                e_name = st.text_input("Person name", value=row["person_name"], key=f"ename_{row['cert_id']}")
                                e_emp = st.text_input("Employee ID", value=row.get("employee_id") or "", key=f"eemp_{row['cert_id']}")
                                e_trade = st.text_input("Trade code", value=row["trade_code"], key=f"etrade_{row['cert_id']}")
                                e_desig = st.text_input("Designation", value=row.get("designation") or "", key=f"edesig_{row['cert_id']}")
                            with e2:
                                e_link = st.text_input("Link", value=row.get("link") or "", key=f"elink_{row['cert_id']}")
                                e_issue = st.date_input(
                                    "Issue date",
                                    value=datetime.strptime(row["issue_date"], "%Y-%m-%d").date(),
                                    key=f"eissue_{row['cert_id']}",
                                )
                                e_expiry = st.date_input(
                                    "Expiry date",
                                    value=datetime.strptime(row["expiry_date"], "%Y-%m-%d").date(),
                                    key=f"eexpiry_{row['cert_id']}",
                                )
                            if st.form_submit_button("Save changes"):
                                update_certificate(
                                    row["cert_id"], e_name, e_emp or None, e_trade.strip().upper() or row["trade_code"],
                                    e_desig, row.get("certificate_no"), e_link or None, row.get("source_path"),
                                    row.get("renamed_path"), row.get("stamped_path"),
                                    e_issue.isoformat(), e_expiry.isoformat(), row["status"],
                                )
                                st.success("Updated.")
                                st.rerun()

            st.divider()
            zc1, zc2 = st.columns(2)
            if zc1.button("📦 Download all certificate QR codes as ZIP"):
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w") as zf:
                    for _, row in df.iterrows():
                        data = build_cert_qr_target(row["cert_id"], cert=row)
                        if not data:
                            continue
                        labeled = common.make_labeled_qr(row["cert_id"], row["person_name"], data)
                        zf.writestr(f"{row['cert_id']}_qr.png", common.img_to_bytes(labeled))
                st.download_button("Click to download ZIP", buf.getvalue(),
                                   file_name="all_certificate_qr_codes.zip", mime="application/zip",
                                   key="dl_zip_certs")
            if zc2.button("📊 Generate Excel (for print vendor)"):
                excel_rows = []
                for _, row in df.iterrows():
                    data = build_cert_qr_target(row["cert_id"], cert=row)
                    if not data:
                        continue
                    excel_rows.append({
                        "cert_id": row["cert_id"], "person_name": row["person_name"],
                        "employee_id": row["employee_id"], "trade_code": row["trade_code"],
                        "designation": row["designation"], "issue_date": row["issue_date"],
                        "expiry_date": row["expiry_date"], "qr_image": common.make_qr_image(data, box_size=8),
                    })
                if not excel_rows:
                    st.warning("No certificates with a generated QR to include (set the App base URL first).")
                else:
                    excel_bytes = common.build_qr_excel(excel_rows, EXCEL_COLUMNS)
                    st.download_button("Click to download Excel", excel_bytes,
                                       file_name="certificates_with_qr.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="dl_xlsx_certs")

    # ── Add Certificate (manual) ──────────────────────────────────────
    with tabs[1]:
        st.subheader("Add a certificate manually")
        st.caption("For one-off entries. For bulk PDFs, use 'Process Certificates' instead.")

        trade_list = get_trades()
        if trade_list:
            trade_for_add = st.selectbox(
                "Trade *", [c for c, _ in trade_list],
                format_func=lambda c: f"{c} — {get_trade_name(c)}", key="add_cert_trade",
            )
        else:
            trade_for_add = st.text_input(
                "Trade code * (no trades set up yet — type one, e.g. CS)", key="add_cert_trade_text",
            ).strip().upper()

        if trade_for_add:
            st.info(f"Certificate ID will be: **{next_cert_id(trade_for_add)}**")

        default_validity = int(common.get_setting(DB_PATH, "dqas_default_validity_months", "6") or 6)
        with st.form("add_cert_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                person_name = st.text_input("Person name *", placeholder="R. Sharma")
                employee_id = st.text_input("Employee ID", placeholder="SJC013989")
                issue_date_in = st.date_input("Issue date *", value=datetime.now().date())
            with c2:
                link = st.text_input("Certificate link (optional, can add later)", placeholder="https://workdrive.zoho.com/file/xxxxx")
                default_expiry = common.add_months(datetime.now().date(), default_validity)
                expiry_date_in = st.date_input("Expiry date *", value=default_expiry)
            submitted = st.form_submit_button("Add certificate", type="primary")

            if submitted:
                if not person_name or not trade_for_add:
                    st.error("Person name and Trade are required.")
                elif expiry_date_in <= issue_date_in:
                    st.error("Expiry date must be after the issue date.")
                else:
                    if not any(c == trade_for_add for c, _ in trade_list):
                        upsert_trade(trade_for_add, trade_for_add)
                    cert_id = next_cert_id(trade_for_add)
                    add_certificate(
                        cert_id, person_name, employee_id or None, trade_for_add, None, None,
                        link or None, None, None, None, issue_date_in.isoformat(), expiry_date_in.isoformat(),
                    )
                    st.success(f"Added {cert_id}. View its QR code in the Certificates tab.")

    # ── Process Certificates ──────────────────────────────────────────
    with tabs[2]:
        st.subheader("Process certificates — extract, rename, QR-generate & stamp automatically")
        st.caption(
            "Point this at a folder of individual certificate PDFs, or at a single combined "
            "multi-page trade-wise PDF. Every certificate is read automatically (name, "
            "employee ID, trade, issue/expiry dates), given a certificate ID in our format, "
            "renamed, and — if the App base URL is set — QR-stamped automatically."
        )

        certs_root = common.get_setting(DB_PATH, "certs_root_folder", DEFAULT_CERTS_ROOT)
        link_mode = common.get_setting(DB_PATH, "link_mode", "static")
        qr_ready, qr_warning = qr_config_status()
        if qr_warning:
            st.warning(
                qr_warning + " Certificates will still be extracted, renamed, and imported, "
                "but QR stamping and the master PDF/Excel will be skipped until then."
            )
        if link_mode == "static":
            st.caption(
                "📄 Link mode: **static verification page**. A `Verify_<Trade>/<random-token>.html` "
                "file is generated per certificate — the filename is a random token, not the "
                "certificate ID, so pages can't be enumerated. Host these as static files (e.g. "
                "GitHub Pages) at the base URL set in Settings."
            )
        elif link_mode == "direct" and not common.zoho_configured(DB_PATH):
            st.info(
                "📄 Link mode: **direct link**, and Zoho WorkDrive isn't configured — QR/stamping "
                "for each certificate will be skipped here since there's no link yet. Add the "
                "link via Edit afterward, then use the 'Stamp Certificates' tab for that certificate."
            )

        source_path = st.text_input(
            "Source folder or PDF file",
            value=certs_root,
            help="A folder of individual certificate PDFs (scanned non-recursively), or one combined multi-page PDF.",
        )

        if st.button("Scan"):
            if not os.path.exists(source_path):
                st.error("That path doesn't exist.")
            else:
                rows, errors = scan_source(source_path)
                st.session_state["proc_rows"] = rows
                st.session_state["proc_errors"] = errors

        if "proc_rows" in st.session_state:
            rows = st.session_state["proc_rows"]
            errors = st.session_state.get("proc_errors", [])
            for e in errors:
                st.warning(e)

            if not rows:
                st.info("No certificates found.")
            else:
                st.success(f"Found {len(rows)} certificate(s). Review below, uncheck any you don't want, fix flagged dates, then process.")
                df_proc = pd.DataFrame(rows).drop(columns=["_file", "_page", "_certificate_no"])
                edited = st.data_editor(
                    df_proc,
                    column_config={"Include": st.column_config.CheckboxColumn()},
                    disabled=["Status"],
                    hide_index=True,
                    key="proc_editor",
                    use_container_width=True,
                )

                if st.button("Process selected certificates", type="primary"):
                    results = []
                    stamped_by_trade = {}
                    excel_rows_by_trade = {}

                    for i, erow in edited.iterrows():
                        if not erow["Include"]:
                            continue

                        issue_s, expiry_s = str(erow["Issue Date"]), str(erow["Expiry Date"])
                        try:
                            datetime.strptime(issue_s, "%Y-%m-%d")
                            datetime.strptime(expiry_s, "%Y-%m-%d")
                        except ValueError:
                            results.append(f"❌ Skipped {erow['Person Name']} — invalid Issue/Expiry date, fix and re-scan.")
                            continue

                        trade_code = (erow["Trade Code"] or "GEN").strip().upper()
                        if not any(c == trade_code for c, _ in get_trades()):
                            upsert_trade(trade_code, erow["Designation"] or trade_code)
                        trade_folder = safe_name(get_trade_name(trade_code))

                        cert_id = next_cert_id(trade_code)
                        base_fname = safe_name(f"{cert_id}_{erow['Person Name']}_{erow['Employee ID'] or 'NA'}")

                        renamed_dir = os.path.join(certs_root, "Renamed", trade_folder)
                        qr_dir = os.path.join(certs_root, f"Certificate with QR_{trade_folder}")
                        os.makedirs(renamed_dir, exist_ok=True)

                        src_row = rows[i]
                        renamed_path = os.path.join(renamed_dir, f"{base_fname}.pdf")
                        try:
                            extract_single_page_pdf(src_row["_file"], src_row["_page"], renamed_path)
                        except Exception as e:
                            results.append(f"❌ {erow['Person Name']}: failed to save renamed file ({e})")
                            continue

                        # "direct" mode needs the certificate's own link *before* the QR
                        # can be built (the QR IS that link) — get it now if possible.
                        link = None
                        if link_mode == "direct" and common.zoho_configured(DB_PATH):
                            link, zerr = common.zoho_upload_and_link(DB_PATH, renamed_path)
                            if zerr:
                                results.append(f"⚠️ {cert_id}: WorkDrive link failed — {zerr}")

                        stamped_path = None
                        if qr_ready:
                            data = build_cert_qr_target(cert_id, cert={"link": link})
                            if data:
                                os.makedirs(qr_dir, exist_ok=True)
                                base_img = common.load_certificate_image(renamed_path)
                                lx = float(common.get_setting(DB_PATH, "stamp_x_pct", "85"))
                                ly = float(common.get_setting(DB_PATH, "stamp_y_pct", "85"))
                                lsize = float(common.get_setting(DB_PATH, "stamp_size_pct", "7.2"))
                                stamped_img = common.stamp_certificate(base_img, cert_id, data, lx, ly, lsize)
                                stamped_img = apply_signature_if_set(stamped_img)
                                stamped_path = os.path.join(qr_dir, f"{base_fname}.pdf")
                                stamped_img.convert("RGB").save(stamped_path, "PDF")
                                stamped_by_trade.setdefault(trade_folder, []).append(stamped_img)
                                excel_rows_by_trade.setdefault(trade_folder, []).append({
                                    "cert_id": cert_id, "person_name": erow["Person Name"],
                                    "employee_id": erow["Employee ID"], "trade_code": trade_code,
                                    "designation": erow["Designation"], "issue_date": issue_s,
                                    "expiry_date": expiry_s, "qr_image": common.make_qr_image(data, box_size=8),
                                })

                        # "static"/"redirect" modes don't need the link before stamping —
                        # get it now (after) if it wasn't already fetched above.
                        if link_mode != "direct" and stamped_path and common.zoho_configured(DB_PATH):
                            link, zerr = common.zoho_upload_and_link(DB_PATH, stamped_path)
                            if zerr:
                                results.append(f"⚠️ {cert_id}: uploaded/renamed OK, but WorkDrive link failed — {zerr}")

                        add_certificate(
                            cert_id, erow["Person Name"], erow["Employee ID"] or None, trade_code,
                            erow["Designation"], src_row["_certificate_no"], link,
                            src_row["_file"], renamed_path, stamped_path, issue_s, expiry_s,
                        )

                        if link_mode == "static" and qr_ready:
                            new_cert = get_certificate(cert_id)
                            verify_dir = os.path.join(certs_root, f"Verify_{trade_folder}")
                            os.makedirs(verify_dir, exist_ok=True)
                            verify_path = os.path.join(verify_dir, f"{new_cert['verify_token']}.html")
                            with open(verify_path, "w", encoding="utf-8") as f:
                                f.write(build_static_verification_html(new_cert))

                        link_note = " + WorkDrive link" if link else ""
                        results.append(f"✅ {cert_id} — {erow['Person Name']} ({trade_code}){link_note}")

                    for trade_folder, imgs in stamped_by_trade.items():
                        master_path = os.path.join(certs_root, f"{trade_folder}_QR_Certificates_Master.pdf")
                        with open(master_path, "wb") as f:
                            f.write(common.images_to_pdf_bytes(imgs))
                        results.append(f"📄 Master PDF saved: {master_path}")

                    for trade_folder, excel_rows in excel_rows_by_trade.items():
                        excel_path = os.path.join(certs_root, f"{trade_folder}_QR_Certificates.xlsx")
                        excel_bytes = common.build_qr_excel(excel_rows, EXCEL_COLUMNS)
                        with open(excel_path, "wb") as f:
                            f.write(excel_bytes)
                        results.append(f"📊 Print Excel saved: {excel_path}")

                    if link_mode == "static" and stamped_by_trade:
                        results.append(
                            "📁 Static verification pages saved in Verify_<Trade>/ folders under "
                            f"{certs_root} — copy/push these to your static host (e.g. the docs/ "
                            "folder of a GitHub Pages site) at the base URL set in Settings."
                        )

                    st.session_state["proc_results"] = results
                    del st.session_state["proc_rows"]
                    st.session_state.pop("proc_errors", None)
                    st.rerun()

        if "proc_results" in st.session_state:
            st.markdown("#### Last run results")
            for r in st.session_state["proc_results"]:
                st.write(r)
            if st.button("Clear results"):
                del st.session_state["proc_results"]
                st.rerun()

    # ── Stamp Certificates (ad-hoc / calibration) ─────────────────────
    with tabs[3]:
        st.subheader("Stamp QR + Certificate ID + GM Signature onto a certificate file")
        st.caption(
            "Fix each position once below — every certificate stamped afterward (including "
            "in 'Process Certificates') uses the same spots. Certificates can be JPG, PNG, or "
            "PDF (the first page is used)."
        )

        sample_file = st.file_uploader(
            "Upload a sample certificate to preview/adjust positions on",
            type=["jpg", "jpeg", "png", "pdf"], key="stamp_sample",
        )
        base_img = common.load_certificate_image(sample_file) if sample_file else None

        st.markdown("#### 1. Fix QR + Certificate ID position")
        cur_x = float(common.get_setting(DB_PATH, "stamp_x_pct", "85"))
        cur_y = float(common.get_setting(DB_PATH, "stamp_y_pct", "85"))
        cur_size = float(common.get_setting(DB_PATH, "stamp_size_pct", "7.2"))
        st.caption(f"Current locked position: X={cur_x:.0f}%, Y={cur_y:.0f}%, Size={cur_size:.0f}%")

        if base_img is not None:
            x_pct = st.slider("Horizontal position (% from left)", 0, 100, int(cur_x), key="stamp_x_slider")
            y_pct = st.slider("Vertical position (% from top)", 0, 100, int(cur_y), key="stamp_y_slider")
            size_pct = st.slider("Stamp size (% of certificate width)", 5, 50, int(cur_size), key="stamp_size_slider")

            preview_data = "https://example.com/preview?cert=CERT-DQAS-CS-01"
            preview_img = common.stamp_certificate(base_img, "CERT-DQAS-CS-01", preview_data, x_pct, y_pct, size_pct)

            if st.button("Lock QR position", type="primary"):
                common.set_setting(DB_PATH, "stamp_x_pct", str(x_pct))
                common.set_setting(DB_PATH, "stamp_y_pct", str(y_pct))
                common.set_setting(DB_PATH, "stamp_size_pct", str(size_pct))
                st.success("QR position locked. Every certificate stamped from now on will use this position.")
        else:
            preview_img = None
            st.info("Upload a sample certificate above to adjust and lock the QR position.")

        st.divider()
        st.markdown("#### 2. GM signature")
        st.caption(
            "Upload once — it's saved and reused for every certificate. A PNG with a "
            "transparent background looks best. Position it above the Authorised "
            "Signatory's printed name."
        )

        sig_upload = st.file_uploader("GM signature (PNG)", type=["png"], key="sig_upload")
        if sig_upload:
            sig_dir = os.path.join(certs_root, "_assets")
            os.makedirs(sig_dir, exist_ok=True)
            sig_path = os.path.join(sig_dir, "gm_signature.png")
            with open(sig_path, "wb") as f:
                f.write(sig_upload.getvalue())
            common.set_setting(DB_PATH, "signature_path", sig_path)
            st.success(f"Signature saved to {sig_path}")
            st.rerun()

        current_sig_path = common.get_setting(DB_PATH, "signature_path", "")
        if not current_sig_path or not os.path.exists(current_sig_path):
            st.info("Upload a GM signature PNG above to enable this step.")
        elif base_img is None:
            st.info("Upload a sample certificate above (in section 1) to preview and lock the signature position.")
        else:
            cur_sx = float(common.get_setting(DB_PATH, "sig_x_pct", "50"))
            cur_sy = float(common.get_setting(DB_PATH, "sig_y_pct", "60"))
            cur_ssize = float(common.get_setting(DB_PATH, "sig_size_pct", "15"))
            st.caption(f"Current locked signature position: X={cur_sx:.0f}%, Y={cur_sy:.0f}%, Size={cur_ssize:.0f}%")

            sig_img = common.load_overlay_image(current_sig_path)
            sx_pct = st.slider("Signature horizontal position (% from left)", 0, 100, int(cur_sx), key="sig_x_slider")
            sy_pct = st.slider("Signature vertical position (% from top)", 0, 100, int(cur_sy), key="sig_y_slider")
            ssize_pct = st.slider("Signature size (% of certificate width)", 3, 40, int(cur_ssize), key="sig_size_slider")

            combined_preview = common.paste_overlay(preview_img if preview_img is not None else base_img, sig_img, sx_pct, sy_pct, ssize_pct)
            st.image(combined_preview, caption="Preview (QR + Signature)", use_container_width=True)

            if st.button("Lock signature position", type="primary"):
                common.set_setting(DB_PATH, "sig_x_pct", str(sx_pct))
                common.set_setting(DB_PATH, "sig_y_pct", str(sy_pct))
                common.set_setting(DB_PATH, "sig_size_pct", str(ssize_pct))
                st.success("Signature position locked. Every certificate stamped from now on will use this position.")

        st.divider()
        st.markdown("#### 3. Stamp a single certificate (ad hoc)")
        df_all_certs = get_certificates()
        if df_all_certs.empty:
            st.info("No certificates yet.")
        else:
            cert_options = (df_all_certs["cert_id"] + " — " + df_all_certs["person_name"]).tolist()
            selected = st.selectbox("Select certificate", cert_options, key="stamp_cert_select")
            sel_cert_id = selected.split(" — ")[0]
            cert = get_certificate(sel_cert_id)

            upload_override = st.file_uploader(
                "Certificate file (auto-loaded from its renamed file if available)",
                type=["jpg", "jpeg", "png", "pdf"], key="stamp_cert_upload",
            )

            source = upload_override
            if source is None:
                for p in (cert.get("renamed_path"), cert.get("source_path")):
                    if p and os.path.exists(p):
                        source = p
                        break

            if source is None:
                st.warning("No certificate file available for this certificate. Upload one above.")
            else:
                data = build_cert_qr_target(sel_cert_id, cert=cert)
                if not data:
                    _, warn_msg = qr_config_status()
                    st.warning(warn_msg or "Set the certificate's Link before generating a QR in direct mode.")
                else:
                    stamp_source_img = common.load_certificate_image(source)
                    stamped = common.stamp_certificate(stamp_source_img, sel_cert_id, data, cur_x, cur_y, cur_size)
                    stamped = apply_signature_if_set(stamped)
                    st.image(stamped, caption=f"Stamped certificate — {sel_cert_id}", use_container_width=True)

                    st.download_button(
                        "Download stamped certificate (PNG)",
                        common.img_to_bytes(stamped),
                        file_name=f"{sel_cert_id}_stamped.png",
                        mime="image/png",
                    )
                    pdf_buf = io.BytesIO()
                    stamped.convert("RGB").save(pdf_buf, format="PDF")
                    st.download_button(
                        "Download stamped certificate (PDF)",
                        pdf_buf.getvalue(),
                        file_name=f"{sel_cert_id}_stamped.pdf",
                        mime="application/pdf",
                    )

                    if common.get_setting(DB_PATH, "link_mode", "static") == "static":
                        verify_html = build_static_verification_html(cert)
                        st.download_button(
                            "Download static verification page (.html)",
                            verify_html.encode("utf-8"),
                            file_name=f"{cert.get('verify_token') or sel_cert_id}.html",
                            mime="text/html",
                        )

    # ── Test ───────────────────────────────────────────────────────────
    with tabs[4]:
        st.subheader("Test a certificate scan")
        st.caption(
            "Preview exactly what a phone scanning this certificate's QR code would show, "
            "as of any date — no need to wait for the real expiry date or deploy a public URL first."
        )
        df_all_certs = get_certificates()
        if df_all_certs.empty:
            st.info("No certificates yet.")
        else:
            cert_options = (df_all_certs["cert_id"] + " — " + df_all_certs["person_name"]).tolist()
            selected = st.selectbox("Select certificate", cert_options, key="test_cert_select")
            sel_cert_id = selected.split(" — ")[0]
            cert = get_certificate(sel_cert_id)

            current_mode = common.get_setting(DB_PATH, "link_mode", "static")
            st.caption(f"Current link mode: **{current_mode}**")

            if current_mode == "static":
                st.markdown(
                    "##### Live preview of the actual static verification page\n"
                    "This renders the real generated page, JavaScript included — its date "
                    "check runs against *today's real date* in your browser (not the "
                    "'as of' date below, which only applies to the app-side preview further "
                    "down)."
                )
                import streamlit.components.v1 as components
                components.html(build_static_verification_html(cert), height=420, scrolling=True)
                st.divider()

            as_of = st.date_input("Simulate scan as of", value=datetime.now().date(), key="test_as_of")
            st.markdown("##### App-side preview (works for any link mode, any date)")
            html = render_verification_html(cert, as_of, redirect=False)
            st.markdown(html, unsafe_allow_html=True)
            if cert.get("link"):
                st.link_button("Open certificate link", cert["link"])

    # ── Manage Trades ──────────────────────────────────────────────────
    with tabs[5]:
        st.subheader("Manage Trades")
        st.caption(
            "The trade code is auto-detected from each certificate (from its Certificate "
            "No., e.g. DQAS-CS-... → CS). Set a friendlier display name here — it's used "
            "for output folder names and shown throughout the app."
        )
        trades = get_trades()
        if trades:
            for code, name in trades:
                c1, c2, c3 = st.columns([1, 3, 1])
                c1.write(f"**{code}**")
                new_name = c2.text_input("Name", value=name, key=f"tname_{code}", label_visibility="collapsed")
                if new_name != name:
                    upsert_trade(code, new_name)
                if c3.button("Delete", key=f"tdel_{code}"):
                    delete_trade(code)
                    st.rerun()
        else:
            st.caption("No trades yet — they'll appear automatically after your first Process run, or add one manually below.")

        with st.form("add_trade_form"):
            c1, c2 = st.columns(2)
            new_code = c1.text_input("Trade code", placeholder="CS")
            new_name = c2.text_input("Display name", placeholder="Concrete Supervisor")
            if st.form_submit_button("Add trade"):
                if new_code.strip():
                    upsert_trade(new_code.strip().upper(), new_name.strip() or new_code.strip().upper())
                    st.success(f"Added: {new_code.strip().upper()}")
                    st.rerun()
                else:
                    st.error("Trade code is required.")

    # ── Settings ───────────────────────────────────────────────────────
    with tabs[6]:
        st.subheader("Settings")
        st.subheader("Link mode")
        st.caption("Controls what each certificate's QR code actually encodes.")

        current_link_mode = common.get_setting(DB_PATH, "link_mode", "static")
        link_mode_options = ["static", "direct", "redirect"]
        link_mode_labels = {
            "static": "Static verification page (no hosting needed)",
            "direct": "Direct link only (simplest, no validity check on scan)",
            "redirect": "App redirect (needs this app deployed at a public URL)",
        }
        new_link_mode = st.radio(
            "QR link mode",
            link_mode_options,
            index=link_mode_options.index(current_link_mode) if current_link_mode in link_mode_options else 0,
            format_func=lambda m: link_mode_labels[m],
        )

        if new_link_mode == "static":
            st.markdown(
                "QR codes link to a small self-contained page (generated per certificate, "
                "see **Process Certificates**) that checks the expiry date **in the visitor's "
                "own browser** — no server needed. Host the generated `Verify_<Trade>/*.html` "
                "files as static files (e.g. GitHub Pages) and set the base URL they'll be "
                "served from below.\n\n"
                "⚠️ Because it's a static page, revoking a certificate *after* its page was "
                "generated won't show up until you regenerate and reshare that page.\n\n"
                "🔒 On a free hosting plan (e.g. GitHub Pages on GitHub Free), these pages are "
                "reachable by anyone with the exact URL even if the repo is private — there's "
                "no real login gate available without a paid plan. Mitigations already built "
                "in: each page's URL is a random per-certificate token, not the certificate ID "
                "(so finding one page doesn't let anyone enumerate the rest), and every page is "
                "marked `noindex` so search engines won't index it. If that's still not enough "
                "for this data, use **Direct link** mode instead — it publishes nothing."
            )
            pages_base_url = st.text_input(
                "Static pages base URL",
                value=common.get_setting(DB_PATH, "pages_base_url", ""),
                placeholder="https://yourusername.github.io/DQAS-Certificate-Manager",
                help="QR codes will encode <this>/<random-token>.html — a per-certificate random token, not the certificate ID, so pages can't be enumerated.",
            )
        elif new_link_mode == "direct":
            st.markdown(
                "QR codes encode the certificate's own link directly — nothing to host, "
                "but no automatic Valid/Expired check on scan. The viewer relies on the "
                "Issue/Valid Until dates already printed on the certificate."
            )
            pages_base_url = common.get_setting(DB_PATH, "pages_base_url", "")
        else:
            st.markdown(
                "QR codes link back to *this app* (`?cert=<id>`), which checks validity "
                "server-side on every scan. Requires deploying this app at a permanent "
                "public URL — the verification endpoint is **publicly accessible**, no "
                "login needed to follow a scan."
            )
            base_url = st.text_input(
                "App base URL",
                value=common.get_setting(DB_PATH, "base_url", ""),
                placeholder="https://yourcompany-dqas.streamlit.app",
            )
            pages_base_url = common.get_setting(DB_PATH, "pages_base_url", "")

        if st.button("Save link mode settings"):
            common.set_setting(DB_PATH, "link_mode", new_link_mode)
            if new_link_mode == "static":
                common.set_setting(DB_PATH, "pages_base_url", pages_base_url)
            elif new_link_mode == "redirect":
                common.set_setting(DB_PATH, "base_url", base_url)
            st.success("Link mode saved.")

        st.divider()
        st.subheader("Certificates root folder")
        st.caption(
            "Used by 'Process Certificates' as the default source location, and as the base "
            "for its output folders (Renamed/, Certificate with QR_<Trade>/, and the master PDFs)."
        )
        certs_root_setting = st.text_input(
            "Certificates root folder",
            value=common.get_setting(DB_PATH, "certs_root_folder", DEFAULT_CERTS_ROOT),
        )
        if st.button("Save certificates root folder"):
            common.set_setting(DB_PATH, "certs_root_folder", certs_root_setting)
            st.success("Saved.")

        st.divider()
        st.subheader("DQAS certificate validity")
        st.caption(
            "Default validity period suggested when adding a certificate manually. "
            "Certificates processed automatically use the Issue/Valid Until dates printed "
            "on the certificate itself, so this setting doesn't affect them."
        )
        current_validity = int(common.get_setting(DB_PATH, "dqas_default_validity_months", "6") or 6)
        new_validity = st.number_input(
            "Default validity (months)", min_value=1, max_value=120,
            value=current_validity, key="dqas_validity_setting",
        )
        if st.button("Save validity setting"):
            common.set_setting(DB_PATH, "dqas_default_validity_months", str(int(new_validity)))
            st.success("Validity setting saved.")

        st.divider()
        st.subheader("Zoho WorkDrive integration")
        st.caption(
            "Lets the app upload certificates to WorkDrive and fetch a public link "
            "automatically — during Process Certificates, and on demand from the "
            "Certificates tab — instead of you copying links in by hand."
        )

        with st.expander("One-time setup: exchange a grant code for a refresh token"):
            st.markdown(
                "From [api-console.zoho.com](https://api-console.zoho.com/): create a "
                "**Self Client**, copy its Client ID/Secret, then under **Generate Code** "
                "request scope `WorkDrive.files.ALL,WorkDrive.team.ALL,WorkDrive.workspace.ALL` "
                "and copy the resulting code here quickly (it expires in minutes, single-use)."
            )
            gc1, gc2 = st.columns(2)
            grant_client_id = gc1.text_input("Client ID", key="zoho_grant_client_id")
            grant_client_secret = gc2.text_input("Client Secret", type="password", key="zoho_grant_client_secret")
            grant_domain = st.text_input(
                "Accounts domain", value=common.get_setting(DB_PATH, "zoho_accounts_domain", "accounts.zoho.in"),
                key="zoho_grant_domain", help="The domain your Zoho account uses, e.g. accounts.zoho.in / .com / .eu",
            )
            grant_code = st.text_input("Grant / authorization code", key="zoho_grant_code")
            if st.button("Exchange code for refresh token"):
                if not (grant_client_id and grant_client_secret and grant_code):
                    st.error("Client ID, Client Secret, and the grant code are all required.")
                else:
                    refresh_token, access_token, err = common.zoho_exchange_grant_code(
                        grant_client_id, grant_client_secret, grant_code, grant_domain,
                    )
                    if err:
                        st.error(err)
                    else:
                        common.set_setting(DB_PATH, "zoho_client_id", grant_client_id)
                        common.set_setting(DB_PATH, "zoho_client_secret", grant_client_secret)
                        common.set_setting(DB_PATH, "zoho_accounts_domain", grant_domain)
                        common.set_setting(DB_PATH, "zoho_refresh_token", refresh_token)
                        st.success("Refresh token obtained and saved — this is permanent, you won't need to redo this step.")
                        st.rerun()

        z1, z2 = st.columns(2)
        with z1:
            zoho_client_id = st.text_input("Client ID", value=common.get_setting(DB_PATH, "zoho_client_id", ""), key="zoho_cid")
            zoho_accounts_domain = st.text_input(
                "Accounts domain", value=common.get_setting(DB_PATH, "zoho_accounts_domain", "accounts.zoho.in"), key="zoho_adomain",
            )
            zoho_workdrive_base = st.text_input(
                "WorkDrive base URL", value=common.get_setting(DB_PATH, "zoho_workdrive_base", "https://workdrive.zoho.in"),
                key="zoho_wdbase", help="e.g. https://workdrive.zoho.in or https://workdrive.zoho.com",
            )
        with z2:
            zoho_client_secret = st.text_input(
                "Client Secret", value=common.get_setting(DB_PATH, "zoho_client_secret", ""), type="password", key="zoho_csec",
            )
            zoho_refresh_token = st.text_input(
                "Refresh token", value=common.get_setting(DB_PATH, "zoho_refresh_token", ""), type="password", key="zoho_rtok",
                help="Filled in automatically by the one-time setup above, or paste one you already have.",
            )
            zoho_folder_id = st.text_input(
                "Target WorkDrive folder ID", value=common.get_setting(DB_PATH, "zoho_parent_folder_id", ""), key="zoho_folder",
                help="From the folder's WorkDrive URL: .../folders/<this part>",
            )

        sc1, sc2 = st.columns(2)
        if sc1.button("Save Zoho settings"):
            common.set_setting(DB_PATH, "zoho_client_id", zoho_client_id)
            common.set_setting(DB_PATH, "zoho_client_secret", zoho_client_secret)
            common.set_setting(DB_PATH, "zoho_accounts_domain", zoho_accounts_domain)
            common.set_setting(DB_PATH, "zoho_workdrive_base", zoho_workdrive_base)
            common.set_setting(DB_PATH, "zoho_refresh_token", zoho_refresh_token)
            common.set_setting(DB_PATH, "zoho_parent_folder_id", zoho_folder_id)
            st.success("Zoho settings saved.")
        if sc2.button("Test Zoho connection"):
            ok, msg = common.zoho_test_connection(DB_PATH)
            (st.success if ok else st.error)(msg)

        st.caption(
            "⚠️ After the first real certificate gets a link, open it in a private/incognito "
            "browser window to confirm it opens without a Zoho login prompt. If it asks you to "
            "log in, tell me and I'll adjust the link's sharing permission."
        )

        st.divider()
        st.subheader("Database backup")
        if os.path.exists(DB_PATH):
            with open(DB_PATH, "rb") as f:
                st.download_button("Download database backup", f.read(), file_name="dqas_certificates_backup.db")
        else:
            st.caption("Database will be created after you add your first certificate.")


main()
