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
                context[k] = d.strftime("%d %B %Y")
            except Exception:
                context[k] = str(v)
        else:
            context[k] = str(v)
    return context


def safe_filename(name: str) -> str:
    """
    Makes a safe filename from contract_id or fallback name.
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
      <body style="font-family: Arial; max-width: 720px; margin: 40px aut
