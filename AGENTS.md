# CT (Contextual Thinking)

## Cursor Cloud specific instructions

This repo has two independent pieces (no package.json, no lint/test framework configured):

1. **`index.html`** — a self-contained interactive slide presentation ("Contextual Thinking"). All JS is inline; Tailwind and Google Fonts load from CDNs, so full styling needs outbound internet. Serve it as static files, e.g. `python3 -m http.server 8000` from the repo root, then open `http://localhost:8000/index.html`. There is no build step.

2. **`extract_table.py`** — a Playwright + openpyxl utility that scrapes the largest table (or ExtJS grid) from a page and writes `extracted data.xlsx`. Run it from the virtualenv:
   `source .venv/bin/activate && python extract_table.py <url>`.
   - It first tries to attach to an already-running Chrome via CDP at `127.0.0.1:9222`, and falls back to launching its own Chromium if that fails.
   - Its real target is an external ASP.NET `OrderExecution.aspx` app that is **not** available in this environment. To verify the script locally, point it at any HTML page that contains a `<table>` (a `file://` URL works).
   - `chromium` for Playwright is installed by the update script; if it goes missing, re-run `playwright install --with-deps chromium`.

### Environment
- Python is externally-managed, so dependencies live in a repo-local `.venv` (created by the update script). Always `source .venv/bin/activate` before running the Python utility.
- The static site itself needs no dependencies.
