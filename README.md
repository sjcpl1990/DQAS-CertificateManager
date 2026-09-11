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
- **Three QR link modes** (Settings → Link mode) — pick whichever fits how
  you're running this:
  - **Static verification page** *(default)* — QR links to a small
    self-contained page, generated per certificate, that checks the expiry
    date **in the visitor's own browser**. No app hosting needed at all —
    just host the generated `.html` files as static files (e.g. GitHub
    Pages, see below). Each page's URL is a random per-certificate token
    (not the certificate ID), and every page is marked `noindex` — so
    finding one page doesn't let anyone enumerate the rest, and search
    engines won't index them. That said: on a free hosting plan there's no
    real login gate available, so anyone with the *exact* link (normally
    obtained only by scanning that specific certificate's QR) can open it.
    If that's not acceptable for this data, use Direct link mode instead —
    it publishes nothing anywhere. Also: revoking a certificate *after* its
    page was generated won't show up until that page is regenerated and
    reshared.
  - **Direct link** — QR encodes the certificate's own link, straight
    through. Nothing to host, but no automatic Valid/Expired check —
    the viewer relies on the dates already printed on the certificate.
  - **App redirect** — QR links back to *this app* (`?cert=<id>`), checked
    server-side on every scan. Requires deploying this app itself at a
    permanent public URL.
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

If your source certificates live somewhere else than the local
`DQAS Certificates/` folder (e.g. a network share), change **Settings →
Certificates root folder** — it's also used as the base for this app's
output folders (`Renamed/`, `Certificate with QR_<Trade>/`,
`Verify_<Trade>/`, master PDFs, and the print Excel files).

### No-hosting setup (static verification page — the default)

1. Enable **GitHub Pages** on this repo: **Settings → Pages → Source: Deploy
   from a branch → Branch: `master`, folder: `/docs`** (a `docs/` folder
   with a placeholder page and a `robots.txt` disallowing crawlers is
   already included, ready to push).
2. Note the URL GitHub gives you, e.g.
   `https://<your-username>.github.io/<repo-name>`.
3. In the app, set **Settings → Link mode → Static verification page →
   Static pages base URL** to that URL (append `/docs` if you didn't set
   `docs/` as the Pages root — check what GitHub shows you).
4. After each **Process Certificates** run, copy the generated
   `Verify_<Trade>/*.html` files into this repo's `docs/` folder, then
   commit and push. GitHub Pages serves them within a minute or two.

**On privacy**: these generated pages carry real names and employee IDs,
and GitHub Pages on the Free plan has no real access-control option — a
private repo's Pages site is still publicly reachable by URL. Two things
are already built in to limit exposure: each page's filename is a random
per-certificate token (not the sequential, guessable certificate ID), and
every page is marked `noindex` so search engines won't index it. That
leaves each page reachable only by whoever has that *specific* link —
normally obtained by scanning that specific certificate's own QR code,
the same exposure the printed certificate already has. If even that's not
acceptable for this data, switch to **Direct link** mode instead, which
publishes nothing anywhere.

### Hosting the app itself (only needed for "App redirect" mode)

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
