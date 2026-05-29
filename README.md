# SlideCraft Lite

A focused fork of [SlideCraft](../slidecraft) with just the essentials. No OCR, no overlays editor, no watermark designer, no filters — just the buttons you actually need to clean a NotebookLM PPTX and ship it.

## Features

- **Upload** a PPTX (single file via button, or drag-drop)
- **Remove Logo** — erases the NotebookLM watermark from every slide
- **Bulk** — process up to **10 PPTX files** at once (combined size ≤ 300 MB), download a ZIP of cleaned decks
- **Undo / Redo** across every destructive op (`Ctrl+Z` / `Ctrl+Y`)
- **Save** — drop a manual checkpoint into the undo chain
- **Reset** — restore every slide from the original uploaded version
- **Present** — fullscreen slideshow (`F5`, navigate with arrow keys)
- **Help** — modal with shortcuts (`H`)
- **Export** — PowerPoint (`.pptx`) or PDF

## Quick start

```bash
# Clone and enter the folder
git clone <your-repo> slidecraft-lite
cd slidecraft-lite

# One-command launcher — creates .venv, installs deps, starts the server
./run.sh              # macOS / Linux
# or run.bat          # Windows cmd.exe
# or .\run.ps1        # Windows PowerShell
```

Open **http://127.0.0.1:5051** in your browser.

### Manual run

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

### LibreOffice (recommended)

Without LibreOffice, PPTX conversion falls back to a low-fidelity Pillow path that only extracts embedded pictures. Install it once:

- **macOS:** `brew install --cask libreoffice`
- **Debian/Ubuntu:** `sudo apt install libreoffice`
- **Windows:** download from <https://libreoffice.org/download>

## Configuration

Environment variables:

| Var | Default | Purpose |
|---|---|---|
| `HOST` | `127.0.0.1` | Bind address. Set to `0.0.0.0` for LAN access (no auth — only on trusted networks). |
| `PORT` | `5051` | HTTP port. |
| `MAX_UPLOAD_MB` | `300` | Upload size cap. Default matches the bulk cap. |
| `FLASK_DEBUG` | `false` | Live reload + verbose tracebacks. |

## Bulk processing limits

- **Files per batch:** 10
- **Combined size:** 300 MB
- For larger jobs, run multiple smaller batches.

## What it doesn't do

If you need any of these, use the full [SlideCraft](../slidecraft) instead:

- OCR / text editing
- Overlays (text boxes, shapes, freehand)
- Watermark designer (add custom watermarks)
- Image filters / color grading
- Crop / rotate
- Templates / version history modal
- Comments / annotations
- Video logo removal
- Image overlays / QR codes

## Project layout

```
slidecraft-lite/
├── app.py                       # Flask app (~500 lines, single file)
├── templates/index.html         # UI markup
├── static/
│   ├── css/app.css              # Styles
│   ├── js/app.js                # Client logic
│   └── slides/                  # Working slide JPGs (gitignored)
│       └── _originals/          # Uploaded originals (for Reset)
├── uploads/                     # Last-uploaded PPTX (gitignored)
├── exports/                     # Downloaded exports (gitignored)
├── history/                     # Auto + manual snapshots (gitignored)
├── ops_log.json                 # Op log for undo/redo (gitignored)
├── requirements.txt
├── run.sh / run.bat / run.ps1   # Launchers
└── README.md
```

## License

Same as the parent project.
