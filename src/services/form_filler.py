"""Service for filling Excel forms with data."""
import json
from openpyxl import Workbook
from typing import Dict
from src.domain.models import FieldDetectionResult, FillResult
from src.domain.exceptions import ExcelProcessingError
from src.utils.cell_helpers import get_writable_cell, is_cell_empty


class FormFiller:
    """Service for filling Excel forms with data."""
    
    def __init__(self, workbook: Workbook):
        self.workbook = workbook
        self.worksheet = workbook.active
    
    def fill(self, detection_result: FieldDetectionResult, data: Dict[str, str]) -> FillResult:
        """Fill Excel workbook with provided data."""
        filled_count = 0
        skipped_count = 0
        errors = []
        
        for label, field in detection_result.fields.items():
            try:
                if label not in data:
                    errors.append(f"'{label}' not found in data")
                    skipped_count += 1
                    continue
                
                value = data[label]
                
                if isinstance(value, (dict, list)):
                    value = json.dumps(value, ensure_ascii=False)
                elif value is None:
                    value = ""
                else:
                    value = str(value)
                
                try:
                    cell = get_writable_cell(
                        field.target_coordinate,
                        self.worksheet,
                        detection_result.merged_map
                    )
                except Exception as e:
                    errors.append(f"Error getting writable cell for '{label}': {e}")
                    skipped_count += 1
                    continue
                
                if not is_cell_empty(cell):
                    skipped_count += 1
                    continue
                
                cell.value = value
                filled_count += 1
                
            except Exception as e:
                errors.append(f"Error filling '{label}': {e}")
                skipped_count += 1
                continue
        
        return FillResult(
            filled_count=filled_count,
            skipped_count=skipped_count,
            errors=errors
        )

