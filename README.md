# DQAS Certificate Manager

A Streamlit app that automatically extracts, renames, QR-stamps, and
verifies DQAS-qualified personnel certificates (Engineers, Supervisors,
Foremen, etc.) — organized by trade, global (no per-site grouping).

## What it does

- **Process Certificates** — point it at a folder of individual certificate
  PDFs, or at a single combined multi-page trade-wise PDF, and it
  automatically:
  1. Extracts each person's Name, Employee ID, Designation, Certificate No.,
     Issue Date, and Valid Until date straight from the certificate's own text
  2. Detects the trade code from the certificate's own Certificate No.
     (e.g. `DQAS-CS-...` → trade `CS`)
  3. Generates a certificate ID in our format (`CERT-DQAS-<TRADE>-01`, `-02`, ...)
  4. Saves a renamed copy into `Renamed\<Trade>\`
  5. Generates its QR code and stamps it onto the certificate automatically —
     plus the GM's signature (a PNG you upload once), both at positions fixed
     once via the **Stamp Certificates** tab — saved into
     `Certificate with QR_<Trade>\`
  6. Compiles all of that batch's stamped certificates into one combined
     `<Trade>_QR_Certificates_Master.pdf`
  7. Generates `<Trade>_QR_Certificates.xlsx` — one row per certificate with
     its QR code embedded as an image, ready to hand to a print vendor
  - A review table lets you uncheck rows, fix flagged/invalid dates, or
    correct the detected trade before confirming — nothing is written until
    you click Process. Re-scanning the same source skips certificates already
    imported.
- QR codes **always** route back through this app (never a direct link to
  the file) — this is what lets validity be checked on every scan. Scanning
  an expired certificate's QR shows a clear "Certificate Validity Expired"
  notice instead of opening the file; a revoked certificate shows
  "Certificate Revoked"; a valid one shows its expiry date and redirects to
  the certificate file.
- **Zoho WorkDrive integration** (optional) — upload certificates and fetch
  a public share link automatically instead of copying links in by hand.
  Configure once under Settings.
- **Test** tab — preview exactly what a scan would show, as of any date,
  without waiting for a real expiry date.
- Download individual or bulk (ZIP) QR codes, or export an Excel sheet
  (certificate details + embedded QR image) for a print vendor.
- Manage trade codes/display names in the Manage Trades tab.

## Run

```
pip install -r requirements.txt
streamlit run app.py
```

This is fully self-contained: on first run it creates its own database and
a default `DQAS Certificates/` certificates-root folder automatically — no
external paths to configure to get started. Both are git-ignored, so a
fresh clone always starts empty; real certificate data never gets committed.

## Required setup

Because certificate QR codes must route through the app to check validity,
you **must** deploy this app at a permanent public URL and set it under
**Settings → App base URL** before generating QR codes.

If your source certificates live somewhere else (e.g. a network share),
change **Settings → Certificates root folder** to point there instead — it's
also used as the base for this app's output folders (`Renamed/`,
`Certificate with QR_<Trade>/`, master PDFs, and the print Excel files).

## Deploying for a permanent public URL

- [Streamlit Community Cloud](https://streamlit.io/cloud) (free, simplest)
- Your own server / internal network with a fixed domain
- Any platform that can run a Python web app (Render, Railway, etc.)

## Data storage

SQLite (`dqas_certificates.db`), created automatically on first run under
your local app-data folder (`%LOCALAPPDATA%\DQASCertificateManager\` on
Windows; `~/DQASCertificateManager/` elsewhere) — **not** next to the
script. This is deliberate: SQLite opens a fresh connection per operation,
which needs real file-locking, and cloud-sync or network drives (OneDrive,
Google Drive, a mapped network share) can silently corrupt the database
file under that pattern ("database disk image is malformed") — a real
failure this project hit during development on a mapped drive. If you want
the database somewhere specific, change `DB_PATH` at the top of `app.py` —
just make sure it points at a genuine local disk, not a synced/network one.

Back the database up regularly; a backup download button is available in
the Settings tab. `.db` files and the local `DQAS Certificates/` folder are
git-ignored — this repo tracks code only, never certificate data.

## Files

- `app.py` — the whole application (UI, database layer, certificate text
  extraction, stamping pipeline)
- `common.py` — shared QR/image rendering, printable sheets, login/auth,
  lookup-table helpers, and the Zoho WorkDrive integration
- `Roboto-Bold.ttf` / `Roboto-Regular.ttf` — fonts used for QR labels and
  printable sheets
