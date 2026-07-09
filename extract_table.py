"""Extract table/grid data from OrderExecution.aspx iframe and write to Excel."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

EXTRACT_JS = """
() => {
  function extractFromTable(table) {
    const headers = [];
    const rows = [];
    const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (headerRow) {
      headerRow.querySelectorAll('th, td').forEach((cell) => {
        headers.push((cell.innerText || cell.textContent || '').trim());
      });
    }
    const bodyRows = table.querySelectorAll('tbody tr');
    const dataRows = bodyRows.length ? bodyRows : table.querySelectorAll('tr');
    dataRows.forEach((row, idx) => {
      if (idx === 0 && !bodyRows.length && headers.length) return;
      const cells = row.querySelectorAll('th, td');
      if (!cells.length) return;
      rows.push(
        Array.from(cells).map((cell) => (cell.innerText || cell.textContent || '').trim())
      );
    });
    return { headers, rows };
  }

  function extractFromExtGrid() {
    if (typeof Ext === 'undefined' || !Ext.ComponentQuery) return null;
    const grids = Ext.ComponentQuery.query('gridpanel, grid');
    if (!grids.length) return null;
    let best = grids[0];
    for (const grid of grids) {
      const count = grid.getStore ? grid.getStore().getCount() : 0;
      const bestCount = best.getStore ? best.getStore().getCount() : 0;
      if (count > bestCount) best = grid;
    }
    const columns = (best.columns || []).filter((col) => col && !col.hidden && col.dataIndex);
    const headers = columns.map((col) => {
      const text = col.text || col.header || col.dataIndex || '';
      return String(text).replace(/<[^>]*>/g, '').trim();
    });
    const rows = [];
    const store = best.getStore();
    if (store) {
      store.each((record) => {
        rows.push(
          columns.map((col) => {
            const value = record.get(col.dataIndex);
            return value == null ? '' : String(value).trim();
          })
        );
      });
    }
    return { headers, rows };
  }

  const extResult = extractFromExtGrid();
  if (extResult && extResult.rows.length) return extResult;

  let bestTable = null;
  let bestScore = 0;
  document.querySelectorAll('table').forEach((table) => {
    const score = table.querySelectorAll('tr').length * table.querySelectorAll('td, th').length;
    if (score > bestScore) {
      bestScore = score;
      bestTable = table;
    }
  });
  if (bestTable) return extractFromTable(bestTable);

  return { headers: [], rows: [] };
}
"""


def write_excel(output_path: Path, headers: list[str], rows: list[list[str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    if headers:
        ws.append(headers)
    for row in rows:
        ws.append(row)

    wb.save(output_path)


def extract_from_page(page, iframe_name: str | None, iframe_url_part: str | None):
    target = page
    if iframe_name or iframe_url_part:
        for frame in page.frames:
            if iframe_name and frame.name == iframe_name:
                target = frame
                break
            if iframe_url_part and iframe_url_part in (frame.url or ""):
                target = frame
                break

    return target.evaluate(EXTRACT_JS)


def find_order_execution_page(browser):
    for context in browser.contexts:
        for page in context.pages:
            url = page.url or ""
            if "OrderExecution.aspx" in url:
                return page
            for frame in page.frames:
                if "OrderExecution.aspx" in (frame.url or ""):
                    return page
    return None


def main() -> int:
    workspace = Path(__file__).resolve().parent
    output_path = workspace / "extracted data.xlsx"
    url = sys.argv[1] if len(sys.argv) > 1 else None
    iframe_name = "Window1_IFrame"

    with sync_playwright() as p:
        browser = None
        page = None

        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            page = find_order_execution_page(browser)
            if page:
                print(f"Connected to existing browser tab: {page.url}")
        except Exception as exc:
            print(f"CDP connect failed ({exc}); falling back to direct navigation.")

        if page is None:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()
            if not url:
                print(
                    "Usage: python extract_table.py <full-url-to-OrderExecution.aspx>\n"
                    "Or start Chrome with: chrome.exe --remote-debugging-port=9222"
                )
                return 1
            print(f"Navigating to {url}")
            page.goto(url, wait_until="networkidle", timeout=120000)
            page.wait_for_timeout(3000)

        result = extract_from_page(page, iframe_name, "OrderExecution.aspx")
        headers = result.get("headers") or []
        rows = result.get("rows") or []

        if not rows:
            result = extract_from_page(page, None, None)
            headers = result.get("headers") or []
            rows = result.get("rows") or []

        if not rows:
            debug_path = workspace / "extract_debug.json"
            debug_path.write_text(json.dumps(result, indent=2), encoding="utf-8")
            print(f"No table rows found. Debug output: {debug_path}")
            return 1

        write_excel(output_path, headers, rows)
        print(f"Wrote {len(rows)} rows to {output_path}")
        if headers:
            print(f"Columns: {', '.join(headers)}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
