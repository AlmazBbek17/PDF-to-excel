import os
import base64
import json
import io
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))


def parse_tables_from_response(text):
    """Parse JSON tables from Claude's response."""
    try:
        start = text.find("[")
        end = text.rfind("]") + 1
        if start != -1 and end > start:
            return json.loads(text[start:end])
    except Exception:
        pass
    return []


def build_excel(tables):
    """Build an Excel workbook from parsed tables."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    header_fill = PatternFill("solid", start_color="2E7D32")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, table in enumerate(tables):
        title = table.get("title", f"Table {i + 1}")[:31]
        sheet_name = title if title else f"Sheet{i + 1}"
        ws = wb.create_sheet(title=sheet_name)

        headers = table.get("headers", [])
        rows = table.get("rows", [])

        if not headers and not rows:
            continue

        # Write title
        if title:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 1))
            cell = ws.cell(row=1, column=1, value=title)
            cell.font = Font(bold=True, name="Arial", size=12)
            cell.alignment = center

        # Write headers
        row_offset = 2 if title else 1
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row_offset, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        # Write data rows
        alt_fill = PatternFill("solid", start_color="F1F8E9")
        for row_idx, row in enumerate(rows, 1):
            fill = alt_fill if row_idx % 2 == 0 else None
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=row_offset + row_idx, column=col_idx, value=value)
                cell.alignment = left
                cell.border = border
                if fill:
                    cell.fill = fill

        # Auto-size columns
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)

        ws.freeze_panes = ws.cell(row=row_offset + 1, column=1)

    if not wb.sheetnames:
        wb.create_sheet("Sheet1")

    return wb


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/extract", methods=["POST"])
def extract():
    data = request.get_json()
    if not data or "image" not in data:
        return jsonify({"error": "No image provided"}), 400

    image_data = data["image"]
    # Strip data URL prefix if present
    if "," in image_data:
        image_data = image_data.split(",", 1)[1]

    media_type = data.get("media_type", "image/png")

    prompt = """You are an expert at extracting tables from PDF document images.

Analyze this PDF page image and extract ALL tables you can find.

Return ONLY a JSON array (no markdown, no extra text) in this exact format:
[
  {
    "title": "Table title or description (empty string if none)",
    "headers": ["Column 1", "Column 2", "Column 3"],
    "rows": [
      ["row1col1", "row1col2", "row1col3"],
      ["row2col1", "row2col2", "row2col3"]
    ]
  }
]

Rules:
- Extract every table on the page, even partial ones
- Preserve all text exactly as shown
- If a cell spans multiple columns, repeat the value or use empty strings
- If no tables found, return: []
- Numbers: preserve formatting (e.g. "$1,234.56", "42%")
- Empty cells: use empty string ""
- Do NOT include any explanation, only the JSON array"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_data,
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    )

    response_text = response.content[0].text.strip()
    tables = parse_tables_from_response(response_text)

    if not tables:
        return jsonify({"error": "No tables found on this page", "raw": response_text}), 422

    wb = build_excel(tables)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="exported_tables.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
