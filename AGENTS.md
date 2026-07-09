# AGENTS.md

## Cursor Cloud specific instructions

This repo has two independent components:

1. **`index.html`** — a self-contained interactive presentation / landing page ("Contextual Thinking"). Tailwind is loaded from a CDN and all JS is inline (10-slide deck engine, an interactive quiz, keyboard/arrow navigation). There is no build step.
2. **`extract_table.py`** — a Playwright + `openpyxl` utility that scrapes an HTML table/ExtJS grid into `extracted data.xlsx`.

### Running the site (primary product)
- Serve statically from the repo root, e.g. `python3 -m http.server 8000`, then open `http://localhost:8000/index.html`. Opening the file directly also works, but a server is preferred because of the `?fullscreen=1` launch path and CDN assets.
- There is no lint/test/build tooling configured for the site; validation is manual (open in a browser and interact with the deck/quiz).

### Running the extractor
- Requires the Playwright Chromium browser (installed by the update script via `playwright install`).
- The script first tries to attach to an existing Chrome via CDP on `127.0.0.1:9222`; if that fails it falls back to launching its own browser and navigating to the URL passed as `argv[1]`.
- It launches with `headless=False`, so a display is required. This VM provides one at `DISPLAY=:1` — run as `DISPLAY=:1 python3 extract_table.py <url>`.
- Output is always written to `extracted data.xlsx` in the repo root (a tracked, initially-empty file). Running the script overwrites it; use `git checkout -- "extracted data.xlsx"` to restore it if you don't intend to commit extracted data.
- Quick smoke test: point it at any local HTML page containing a `<table>` (e.g. a `file://` URL) and confirm it prints the row count and column headers.

### Notes
- `pip install --user` puts the `playwright` CLI in `~/.local/bin` (not on PATH); invoke it as `python3 -m playwright ...`.
- `template.pptx`, `content.txt`, and `assets/` are source content for the presentation, not code.
