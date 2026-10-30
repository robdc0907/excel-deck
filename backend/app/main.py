from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import shutil
import os

from backend.app.services.excel_service import ExcelService
from backend.app.services.ppt_service import PPTService

BASE_DIR = Path(__file__).resolve().parent.parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Excel Deck API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

excel_service = ExcelService()
ppt_service = PPTService()

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "frontend")), name="static")


@app.get("/api/health")
def health_check():
    return {"status": "ok", "message": "Excel Deck API is running"}


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file selected")

    safe_name = file.filename.replace(" ", "_")
    file_path = UPLOAD_DIR / safe_name

    with file_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        workbook = excel_service.load_workbook(file_path)
        metadata = excel_service.get_sheet_metadata(workbook)
        return {"filename": safe_name, "sheets": metadata}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Unable to read file: {str(exc)}")


@app.post("/api/generate-ppt")
async def generate_ppt(payload: dict):
    try:
        file_name = payload.get("fileName")
        sheet_name = payload.get("sheetName")
        columns = payload.get("columns", [])
        title = payload.get("title", "Executive Summary")

        if not file_name or not sheet_name or not columns:
            raise HTTPException(status_code=400, detail="Missing fileName, sheetName or columns")

        file_path = UPLOAD_DIR / file_name
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="Uploaded file not found")

        workbook = excel_service.load_workbook(file_path)
        sheet = workbook[sheet_name]
        rows = excel_service.read_rows(sheet)

        selected_columns = []
        for column in columns:
            column_name = column.get("name")
            formula = column.get("formula", "direct")
            if not column_name:
                continue
            selected_columns.append({"name": column_name, "formula": formula})

        output_path = ppt_service.generate_presentation(
            rows=rows,
            selected_columns=selected_columns,
            title=title,
            file_name=file_name,
        )

        return FileResponse(
            path=str(output_path),
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/")
def serve_index():
    file_path = BASE_DIR / "frontend" / "index.html"
    return FileResponse(file_path)
