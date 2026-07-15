from __future__ import annotations

import csv
import html
import io
import os
import re
from collections import defaultdict
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Iterable

import xlrd
from openpyxl import Workbook, load_workbook


HOST = "127.0.0.1"
PORT = 8000
ASSET_TEMPLATE = os.path.join(os.path.dirname(__file__), "assets", "inventory_export_1.csv")
TARGET_LOCATION = "Furniture of America"
NOT_STOCKED = "not stocked"

SKU_KEYWORDS = ("sku", "itemcode", "item code", "item", "product code", "product")
QTY_EXACT = ("qty", "quantity", "qty now", "stock qty", "available qty", "on hand qty")
QTY_PARTIAL = ("qty", "quantity", "stock", "on hand", "available")


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def clean_sku(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text


def parse_quantity(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def score_sku_header(header: str) -> int:
    if not header:
        return 0
    if header in SKU_KEYWORDS:
        return 100
    if any(keyword in header for keyword in SKU_KEYWORDS):
        return 50
    return 0


def score_qty_header(header: str) -> int:
    if not header:
        return 0
    if header in QTY_EXACT:
        return 100
    if "qty now" in header:
        return 95
    if any(keyword in header for keyword in QTY_PARTIAL):
        score = 60
        if "w" in header and "qty" in header:
            score -= 20
        if "coming" in header or "week" in header or "eta" in header:
            score -= 25
        return score
    return 0


def build_combined_headers(rows: list[list[object]], header_end_index: int) -> list[str]:
    max_cols = max((len(row) for row in rows[: header_end_index + 1]), default=0)
    combined_headers: list[str] = []

    for col_index in range(max_cols):
        pieces: list[str] = []
        for row_index in range(header_end_index + 1):
            row = rows[row_index]
            value = normalize_text(row[col_index] if col_index < len(row) else "")
            if value and value not in pieces:
                pieces.append(value)
        combined_headers.append(" ".join(pieces).strip())

    return combined_headers


def header_score(row: list[object]) -> int:
    row_score = 0
    for cell in row:
        text = normalize_text(cell)
        row_score += score_sku_header(text)
        row_score += score_qty_header(text)
    return row_score


def detect_header_row(rows: list[list[object]]) -> int:
    best_index = 0
    best_score = -1

    limit = min(len(rows), 12)
    for row_index in range(limit):
        row_score = header_score(rows[row_index])
        if row_score > best_score:
            best_score = row_score
            best_index = row_index

    return best_index


def detect_header_block_end(rows: list[list[object]]) -> int:
    limit = min(len(rows), 12)
    scored_rows = [index for index in range(limit) if header_score(rows[index]) > 0]
    if scored_rows:
        return scored_rows[-1]
    return detect_header_row(rows)


def detect_columns(headers: list[str]) -> tuple[int, int]:
    sku_index = -1
    sku_score = -1
    qty_index = -1
    qty_score = -1

    for index, header in enumerate(headers):
        current_sku_score = score_sku_header(header)
        current_qty_score = score_qty_header(header)

        if current_sku_score > sku_score:
            sku_score = current_sku_score
            sku_index = index

        if current_qty_score > qty_score:
            qty_score = current_qty_score
            qty_index = index

    if sku_score <= 0:
        raise ValueError("Could not find an SKU column.")
    if qty_score <= 0:
        raise ValueError("Could not find a quantity column.")

    return sku_index, qty_index


def aggregate_rows(rows: Iterable[list[object]], sku_index: int, qty_index: int) -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)

    for row in rows:
        if sku_index >= len(row):
            continue

        sku = clean_sku(row[sku_index])
        if not sku:
            continue

        qty_value = row[qty_index] if qty_index < len(row) else None
        qty = parse_quantity(qty_value)
        totals[sku] += qty

    return totals


def read_csv_inventory(file_bytes: bytes) -> dict[str, float]:
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            decoded = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV file encoding is not supported.")

    all_rows = list(csv.reader(io.StringIO(decoded)))
    if not all_rows:
        return {}

    header_end_index = detect_header_block_end(all_rows)
    headers = build_combined_headers(all_rows[: header_end_index + 1], header_end_index)
    sku_index, qty_index = detect_columns(headers)
    data_rows = all_rows[header_end_index + 1 :]
    return aggregate_rows(data_rows, sku_index, qty_index)


def read_xlsx_inventory(file_bytes: bytes) -> dict[str, float]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    for worksheet in workbook.worksheets:
        preview_rows: list[list[object]] = []
        all_rows: list[list[object]] = []

        for row in worksheet.iter_rows(values_only=True):
            row_values = list(row)
            all_rows.append(row_values)
            if len(preview_rows) < 12:
                preview_rows.append(row_values)

        if not any(any(cell is not None and str(cell).strip() for cell in row) for row in all_rows):
            continue

        header_end_index = detect_header_block_end(preview_rows)
        headers = build_combined_headers(preview_rows, header_end_index)
        sku_index, qty_index = detect_columns(headers)
        data_rows = all_rows[header_end_index + 1 :]
        return aggregate_rows(data_rows, sku_index, qty_index)

    return {}


def read_xls_inventory(file_bytes: bytes) -> dict[str, float]:
    workbook = xlrd.open_workbook(file_contents=file_bytes)

    for sheet in workbook.sheets():
        all_rows = [sheet.row_values(row_index) for row_index in range(sheet.nrows)]
        if not all_rows:
            continue

        preview_rows = all_rows[:12]
        header_end_index = detect_header_block_end(preview_rows)
        headers = build_combined_headers(preview_rows, header_end_index)
        sku_index, qty_index = detect_columns(headers)
        data_rows = all_rows[header_end_index + 1 :]
        return aggregate_rows(data_rows, sku_index, qty_index)

    return {}


def read_inventory_file(filename: str, file_bytes: bytes) -> dict[str, float]:
    extension = os.path.splitext(filename.lower())[1]

    if extension == ".csv":
        return read_csv_inventory(file_bytes)
    if extension == ".xlsx":
        return read_xlsx_inventory(file_bytes)
    if extension == ".xls":
        return read_xls_inventory(file_bytes)

    raise ValueError(f"Unsupported file type: {extension or filename}")


def merge_inventory(files: list[tuple[str, bytes]]) -> dict[str, float]:
    grand_totals: dict[str, float] = defaultdict(float)

    for filename, file_bytes in files:
        file_totals = read_inventory_file(filename, file_bytes)
        for sku, qty in file_totals.items():
            grand_totals[sku] += qty

    return dict(sorted(grand_totals.items()))


def format_qty_for_export(qty: float) -> str:
    if float(qty).is_integer():
        return str(int(qty))
    return str(qty)


def decode_csv_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV file encoding is not supported.")


def load_asset_rows(asset_file: tuple[str, bytes] | None) -> tuple[list[dict[str, str]], list[str]]:
    if asset_file is None:
        if not os.path.exists(ASSET_TEMPLATE):
            raise ValueError(f"Asset template not found: {ASSET_TEMPLATE}")
        with open(ASSET_TEMPLATE, newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            rows = list(reader)
            fieldnames = reader.fieldnames
    else:
        filename, file_bytes = asset_file
        if os.path.splitext(filename.lower())[1] != ".csv":
            raise ValueError("The asset export file must be a CSV.")
        reader = csv.DictReader(io.StringIO(decode_csv_bytes(file_bytes)))
        rows = list(reader)
        fieldnames = reader.fieldnames

    if not fieldnames:
        raise ValueError("Asset template has no header row.")

    return rows, fieldnames


def build_updated_inventory_csv(
    totals: dict[str, float],
    asset_file: tuple[str, bytes] | None = None,
) -> bytes:
    rows, fieldnames = load_asset_rows(asset_file)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=fieldnames, lineterminator="\n")
    writer.writeheader()

    for row in rows:
        location = clean_sku(row.get("Location"))
        sku = clean_sku(row.get("SKU"))
        if location == TARGET_LOCATION:
            available_value = normalize_text(row.get("Available (not editable)"))
            if available_value != NOT_STOCKED:
                row["On hand (new)"] = format_qty_for_export(totals.get(sku, 0.0))
        writer.writerow(row)

    return output.getvalue().encode("utf-8-sig")


def build_summary_workbook(totals: dict[str, float]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "SKU Totals"
    worksheet.append(["SKU", "Total Qty"])

    for sku, qty in totals.items():
        if float(qty).is_integer():
            worksheet.append([sku, int(qty)])
        else:
            worksheet.append([sku, qty])

    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 14

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.read()


def render_home(message: str = "") -> bytes:
    escaped_message = html.escape(message)
    html_body = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Inventory Process Hub</title>
  <style>
    :root {{
      --bg: #f6f7f4;
      --panel: rgba(255, 255, 255, 0.82);
      --text: #17201b;
      --muted: #66736c;
      --line: rgba(23, 32, 27, 0.08);
      --accent: #163f31;
      --accent-soft: #edf3ef;
      --error-bg: #fff3f2;
      --error-text: #9a3b35;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(22, 63, 49, 0.08), transparent 24%),
        radial-gradient(circle at bottom right, rgba(22, 63, 49, 0.06), transparent 20%),
        linear-gradient(180deg, #f8faf8 0%, #f3f5f2 100%);
      color: var(--text);
    }}
    .wrap {{
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 32px 18px;
    }}
    .panel {{
      width: min(720px, 100%);
      background: var(--panel);
      border: 1px solid rgba(255,255,255,0.65);
      border-radius: 28px;
      padding: 34px;
      backdrop-filter: blur(14px);
      box-shadow: 0 18px 50px rgba(28, 42, 35, 0.08);
    }}
    .label {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 12px;
      border-radius: 999px;
      background: var(--accent-soft);
      color: var(--accent);
      font-size: 0.78rem;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }}
    h1 {{
      margin: 18px 0 10px;
      font-size: clamp(2rem, 5vw, 3rem);
      line-height: 1.02;
      letter-spacing: -0.04em;
    }}
    .intro {{
      margin: 0 0 24px;
      color: var(--muted);
      font-size: 1rem;
      line-height: 1.65;
    }}
    .note {{
      margin-bottom: 18px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.7);
      color: var(--muted);
      line-height: 1.55;
    }}
    .error {{
      background: var(--error-bg);
      color: var(--error-text);
    }}
    form {{
      display: grid;
      gap: 16px;
    }}
    .upload-zone {{
      border: 1px dashed rgba(23, 32, 27, 0.12);
      background: rgba(255, 255, 255, 0.72);
      border-radius: 22px;
      padding: 20px;
    }}
    .upload-meta {{
      display: grid;
      gap: 4px;
      margin-bottom: 12px;
    }}
    .upload-meta strong {{
      font-size: 1rem;
    }}
    .upload-meta span {{
      color: var(--muted);
      font-size: 0.94rem;
    }}
    input[type=file] {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 16px;
      background: white;
      font-size: 0.96rem;
    }}
    button {{
      border: 0;
      border-radius: 16px;
      padding: 15px 20px;
      background: var(--accent);
      color: white;
      font-size: 0.98rem;
      font-weight: 600;
      cursor: pointer;
      box-shadow: 0 10px 24px rgba(22, 63, 49, 0.16);
      transition: transform 160ms ease, box-shadow 160ms ease;
    }}
    button:hover {{
      transform: translateY(-1px);
      box-shadow: 0 14px 28px rgba(22, 63, 49, 0.2);
    }}
    .meta {{
      margin-top: 18px;
      padding-top: 16px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.55;
    }}
    @media (max-width: 720px) {{
      .wrap {{
        padding: 18px 14px;
      }}
      .panel {{
        padding: 24px;
        border-radius: 22px;
      }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="panel">
      <span class="label">Inventory Process</span>
      <h1>Update stock export</h1>
      <p class="intro">Upload your report files and generate the final CSV for Furniture of America. Mixed <strong>CSV</strong>, <strong>XLSX</strong>, and <strong>XLS</strong> files are supported.</p>
      <div class="note {('error' if message else '')}">{escaped_message or 'Stocked SKUs are updated from merged totals. Rows marked not stocked stay blank.'}</div>
      <form method="post" enctype="multipart/form-data">
        <div class="upload-zone">
          <div class="upload-meta">
            <strong>Source files</strong>
            <span>Select two or more inventory reports.</span>
          </div>
          <input type="file" name="inventory_files" accept=".csv,.xlsx,.xls" multiple required>
        </div>
        <div class="upload-zone">
          <div class="upload-meta">
            <strong>Asset export file</strong>
            <span>Optional. Upload a different export CSV if you want to replace the default base file.</span>
          </div>
          <input type="file" name="asset_file" accept=".csv">
        </div>
        <button type="submit">Generate Updated Export</button>
      </form>
      <div class="meta">Default base file: <strong>assets/inventory_export_1.csv</strong></div>
    </div>
  </div>
</body>
</html>"""
    return html_body.encode("utf-8")


def parse_uploaded_files(
    handler: BaseHTTPRequestHandler,
) -> tuple[list[tuple[str, bytes]], tuple[str, bytes] | None]:
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("Expected a multipart form upload.")

    content_length = int(handler.headers.get("Content-Length", "0"))
    if content_length <= 0:
        raise ValueError("No uploaded file data was received.")

    body = handler.rfile.read(content_length)
    message = BytesParser(policy=default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )

    files: list[tuple[str, bytes]] = []
    asset_file: tuple[str, bytes] | None = None
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        field_name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        if not filename:
            continue

        content = part.get_payload(decode=True) or b""
        if not content:
            continue

        if field_name == "inventory_files":
            files.append((filename, content))
        elif field_name == "asset_file":
            asset_file = (filename, content)

    return files, asset_file


class InventoryHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        body = render_home()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        try:
            files, asset_file = parse_uploaded_files(self)

            if len(files) < 2:
                raise ValueError("Please upload at least two inventory files.")

            totals = merge_inventory(files)
            if not totals:
                raise ValueError("No SKU records were found in the uploaded files.")

            output_bytes = build_updated_inventory_csv(totals, asset_file=asset_file)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"inventory_export_updated_{timestamp}.csv"

            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "text/csv; charset=utf-8",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
            self.send_header("Content-Length", str(len(output_bytes)))
            self.end_headers()
            self.wfile.write(output_bytes)
        except Exception as exc:
            body = render_home(str(exc))
            self.send_response(HTTPStatus.BAD_REQUEST)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def log_message(self, format: str, *args: object) -> None:
        return


def run_server() -> None:
    server = ThreadingHTTPServer((HOST, PORT), InventoryHandler)
    print(f"Inventory app running at http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    run_server()
