import os
import json
import base64
import io
import anthropic
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

SYSTEM_PROMPT = """You are a PDF table extraction expert. 
You receive raw text extracted from a PDF page with positional data.
Your job is to identify tables and return them as structured JSON.

Rules:
- Find ALL tables in the text
- For each table return: title (if visible), headers (array), rows (array of arrays)
- If multiple tables have the same structure (same headers), mark them with same_structure: true
- Suggest sheet organization: if tables have different structures → different sheets; if same structure → same sheet stacked
- Return ONLY valid JSON, no markdown, no explanation

Response format:
{
  "tables": [
    {
      "id": 1,
      "title": "Table title or null",
      "page": 1,
      "headers": ["Col1", "Col2", "Col3"],
      "rows": [["val1", "val2", "val3"], ...],
      "same_structure_group": "A"
    }
  ],
  "sheet_suggestion": [
    {
      "sheet_name": "Sheet 1",
      "table_ids": [1, 2],
      "reason": "Same structure, stack vertically"
    }
  ]
}"""


def extract_text_from_pdf(pdf_bytes):
    """Extract text and tables from PDF using pdfplumber"""
    pages_data = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []
            pages_data.append({
                "page": i + 1,
                "text": text,
                "raw_tables": tables
            })
    return pages_data


def analyze_with_claude(pages_data, mode="all_tables"):
    """Send extracted text to Claude for intelligent table analysis"""
    
    # Build prompt based on mode
    content_parts = []
    
    for page in pages_data:
        content_parts.append(f"=== PAGE {page['page']} ===")
        if page['text']:
            content_parts.append(page['text'])
        if page['raw_tables']:
            content_parts.append(f"[pdfplumber detected {len(page['raw_tables'])} table(s) on this page]")
    
    full_text = "\n".join(content_parts)
    
    if mode == "all_tables":
        user_msg = f"Extract all tables from this PDF content and organize them optimally:\n\n{full_text}"
    else:
        user_msg = f"Extract all tables including text context from this PDF:\n\n{full_text}"
    
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}]
    )
    
    raw = response.content[0].text.strip()
    # Clean up if Claude added markdown fences
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


def build_excel(analysis_result):
    """Build Excel file from Claude's analysis"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    tables_by_id = {t["id"]: t for t in analysis_result.get("tables", [])}
    suggestions = analysis_result.get("sheet_suggestion", [])
    
    # If no suggestions, put everything on one sheet
    if not suggestions:
        suggestions = [{
            "sheet_name": "Tables",
            "table_ids": [t["id"] for t in analysis_result.get("tables", [])],
            "reason": "All tables"
        }]
    
    for suggestion in suggestions:
        sheet_name = suggestion["sheet_name"][:31]  # Excel limit
        ws = wb.create_sheet(title=sheet_name)
        current_row = 1
        
        for table_id in suggestion["table_ids"]:
            table = tables_by_id.get(table_id)
            if not table:
                continue
            
            # Add table title if exists
            if table.get("title"):
                title_cell = ws.cell(row=current_row, column=1, value=table["title"])
                title_cell.font = Font(bold=True, size=12)
                current_row += 1
            
            headers = table.get("headers", [])
            rows = table.get("rows", [])
            
            if not headers and not rows:
                continue
            
            # Write headers
            if headers:
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_align
                current_row += 1
            
            # Write data rows
            for row_data in rows:
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=current_row, column=col_idx, value=value)
                current_row += 1
            
            current_row += 2  # Gap between tables
        
        # Auto-fit columns
        for column in ws.columns:
            max_length = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
    
    if not wb.sheetnames:
        wb.create_sheet("Empty")
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert():
    """Main conversion endpoint"""
    try:
        data = request.get_json()
        if not data or "pdf_base64" not in data:
            return jsonify({"error": "Missing pdf_base64"}), 400
        
        mode = data.get("mode", "all_tables")  # all_tables | all_content
        
        # Decode PDF
        pdf_bytes = base64.b64decode(data["pdf_base64"])
        
        # Extract text
        pages_data = extract_text_from_pdf(pdf_bytes)
        
        # Analyze with Claude
        analysis = analyze_with_claude(pages_data, mode=mode)
        
        # Build Excel
        excel_buffer = build_excel(analysis)
        
        # Return as base64
        excel_b64 = base64.b64encode(excel_buffer.read()).decode("utf-8")
        
        return jsonify({
            "success": True,
            "excel_base64": excel_b64,
            "tables_found": len(analysis.get("tables", [])),
            "sheets": [s["sheet_name"] for s in analysis.get("sheet_suggestion", [])],
            "analysis_summary": analysis.get("sheet_suggestion", [])
        })
    
    except json.JSONDecodeError as e:
        return jsonify({"error": f"Claude response parse error: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/preview", methods=["POST"])
def preview():
    """Return table structure without generating Excel - for UI preview"""
    try:
        data = request.get_json()
        if not data or "pdf_base64" not in data:
            return jsonify({"error": "Missing pdf_base64"}), 400
        
        pdf_bytes = base64.b64decode(data["pdf_base64"])
        pages_data = extract_text_from_pdf(pdf_bytes)
        analysis = analyze_with_claude(pages_data)
        
        # Return lightweight preview (no row data)
        preview_tables = []
        for t in analysis.get("tables", []):
            preview_tables.append({
                "id": t["id"],
                "title": t.get("title"),
                "page": t.get("page"),
                "headers": t.get("headers", []),
                "row_count": len(t.get("rows", [])),
                "same_structure_group": t.get("same_structure_group")
            })
        
        return jsonify({
            "success": True,
            "tables": preview_tables,
            "sheet_suggestion": analysis.get("sheet_suggestion", [])
        })
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
