"""API dependencies for dependency injection."""
from openpyxl import load_workbook
from io import BytesIO
from fastapi import UploadFile, HTTPException
from src.services.ai_generator import AIGenerator
from src.config import config


def validate_excel_file(filename: str) -> None:
    """Validate that the uploaded file is an Excel file."""
    if not filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400,
            detail="File must be an Excel file (.xlsx or .xls)"
        )


async def get_workbook_from_upload(file: UploadFile) -> tuple:
    """Create workbook from uploaded file."""
    validate_excel_file(file.filename)
    contents = await file.read()
    workbook = load_workbook(BytesIO(contents))
    return workbook, file.filename


def get_ai_generator() -> AIGenerator:
    """Create AI generator instance."""
    return AIGenerator()

