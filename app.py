import io
import re
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from dateutil.parser import parse as parse_date
from docxtpl import DocxTemplate
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse

# -------------------------------
# App initialization
# -------------------------------
app = FastAPI(title="Contract Drafting Tool")

# Columns that should be treated as dates
DATE_COLS = {"effective_date", "start_date", "end_date"}

# -------------------------------
# Helper functions
# -------------------------------

def normalize_row(row: dict) -> dict:
    """
    Cleans and formats one Excel row into a context dictionary
    suitable for docxtpl rendering.
    """
    context = {}
    for k, v in row.items():
        if pd.isna(v):
            context[k] = ""
        elif k in DATE_COLS:
            try:
                d = parse_date(str(v), dayfirst=True)
                context[k] = d.strftime("%d %B %Y")  # e.g., 15 December 2025
            except Exception:
                # If date parsing fails, keep raw value
                context[k] = str(v)
        else:
            context[k] = str(v)
    return context


def safe_filename(name: str) -> str:
    """
    Makes a safe filename from contract_id or fallback name.
    Keeps only letters, numbers, underscore, hyphen. Truncates length.
    """
    name = name.strip() if name else "contract"
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", name)
    return name[:80]


# -------------------------------
# Web UI (Home Page)
# -------------------------------

@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
      <head>
        <title>Contract Drafting Tool</title>
      </head>
      <body style="font-family: Arial; max-width: 720px; margin: 40px auto;">
        <h2>Contract Drafting Tool (DOCX)</h2>

        <p>
          Upload a Word (.docx) template containing <b>{{placeholders}}</b>
          and an Excel (.xlsx) file where column headers match the placeholders.
        </p>

        <form action="/generate" method="post" enctype="multipart/form-data">

          <div style="margin: 12px 0;">
            <label><b>DOCX Template</b></label><br/>
            <input type="file" name="template" accept=".docx" required>
          </div>

          <div style="margin: 12px 0;">
            <label><b>Excel Input</b></label><br/>
            <input type="file" name="excel" accept=".xlsx" required>
          </div>

          <button type="submit" style="padding: 8px 16px;">
            Generate Contracts
          </button>

        </form>

        <hr/>

        <p style="font-size: 0.9em;">
          <b>Notes:</b><br/>
          • One Excel row = one contract<br/>
          • Excel column names must exactly match DOCX placeholders<br/>
          • Output will be a ZIP file of Word documents
        </p>
      </body>
    </html>
    """


# -------------------------------
# Contract generation endpoint
# -------------------------------

@app.post("/generate")
async def generate_contracts(
    template: UploadFile = File(...),
    excel: UploadFile = File(...)
):
    # Validate file types
    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Template must be a .docx file")
    if not excel.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Excel file must be a .xlsx file")

    with TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        template_path = tmpdir / "template.docx"
        excel_path = tmpdir / "input.xlsx"
        output_dir = tmpdir / "outputs"
        output_dir.mkdir(exist_ok=True)

        # Save uploads to temp files
        template_path.write_bytes(await template.read())
        excel_path.write_bytes(await excel.read())

        # Read Excel
        df = pd.read_excel(excel_path, engine="openpyxl")
        if df.empty:
            raise HTTPException(status_code=400, detail="Excel file has no data rows")

        generated_count = 0

        # Generate DOCX per row
        for idx, row in df.iterrows():
            context = normalize_row(row.to_dict())

            doc = DocxTemplate(str(template_path))
            doc.render(context)

            contract_id = context.get("contract_id", f"contract_{idx + 1}")
            filename = f"{safe_filename(contract_id)}.docx"
            out_file = output_dir / filename
            doc.save(str(out_file))

            generated_count += 1

        # Create ZIP in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file_path in output_dir.glob("*.docx"):
                zipf.write(file_path, arcname=file_path.name)

        zip_buffer.seek(0)

        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={
                "Content-Disposition": f'attachment; filename="contracts_{generated_count}.zip"'
            }
        )    
