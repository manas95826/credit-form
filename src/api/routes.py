"""API route handlers."""
import json
from io import BytesIO
from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from fastapi.responses import StreamingResponse
from typing import Optional

from src.services.excel_processor import ExcelProcessor
from src.services.ai_generator import AIGenerator
from src.services.form_filler import FormFiller
from src.domain.exceptions import AIGenerationError, FieldDetectionError, ExcelProcessingError
from src.api.dependencies import (
    get_workbook_from_upload,
    get_ai_generator,
)

router = APIRouter()


@router.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "message": "Excel Form Filler API",
        "description": "Upload an Excel file, automatically detect form fields, and fill them with AI-generated data",
        "main_endpoint": {
            "url": "/process",
            "method": "POST",
            "description": "Upload Excel file → Detect fields → Fill with AI data → Return filled file"
        },
        "other_endpoints": {
            "/detect": "POST - Detect form fields in Excel file (returns field list)",
            "/fill": "POST - Fill Excel file with AI-generated or custom data"
        }
    }


@router.post("/process")
async def process_excel(
    file: UploadFile = File(...),
    ai_generator: AIGenerator = Depends(get_ai_generator),
):
    """
    Main endpoint - Complete workflow: Upload Excel → Detect fields → Fill with AI → Return filled file.
    """
    try:
        workbook, filename = await get_workbook_from_upload(file)
        
        processor = ExcelProcessor(workbook)
        detection_result = processor.detect_fields()
        
        if not detection_result.fields:
            raise HTTPException(
                status_code=400,
                detail=(
                    "No form fields detected in the file. Make sure your Excel file contains "
                    "form labels (text ending with ':' or containing form keywords)."
                )
            )
        
        field_labels = list(detection_result.fields.keys())
        data = ai_generator.generate_data(field_labels)
        
        filler = FormFiller(workbook)
        fill_result = filler.fill(detection_result, data)
        
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=filled_{filename}",
                "X-Fields-Detected": str(detection_result.count),
                "X-Fields-Filled": str(fill_result.filled_count),
                "X-Fields-Skipped": str(fill_result.skipped_count)
            }
        )
    except (AIGenerationError, FieldDetectionError, ExcelProcessingError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


@router.post("/detect")
async def detect_form_fields(file: UploadFile = File(...)):
    """Detect form fields in an uploaded Excel file."""
    try:
        workbook, filename = await get_workbook_from_upload(file)
        processor = ExcelProcessor(workbook)
        detection_result = processor.detect_fields()
        
        return {
            "filename": filename,
            "fields_detected": detection_result.count,
            "fields": {
                label: field.target_coordinate
                for label, field in detection_result.fields.items()
            }
        }
    except FieldDetectionError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


@router.post("/fill")
async def fill_form(
    file: UploadFile = File(...),
    use_ai: bool = True,
    custom_data: Optional[str] = None,
    ai_generator: AIGenerator = Depends(get_ai_generator),
):
    """
    Fill Excel file with data (advanced endpoint with custom data option).
    """
    try:
        workbook, filename = await get_workbook_from_upload(file)
        
        processor = ExcelProcessor(workbook)
        detection_result = processor.detect_fields()
        
        if not detection_result.fields:
            raise HTTPException(
                status_code=400,
                detail="No form fields detected in the file"
            )
        
        if use_ai:
            field_labels = list(detection_result.fields.keys())
            data = ai_generator.generate_data(field_labels)
        else:
            if not custom_data:
                raise HTTPException(
                    status_code=400,
                    detail="custom_data required when use_ai is False"
                )
            data = json.loads(custom_data)
        
        filler = FormFiller(workbook)
        fill_result = filler.fill(detection_result, data)
        
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=filled_{filename}"}
        )
    except (AIGenerationError, FieldDetectionError, ExcelProcessingError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON in custom_data: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

