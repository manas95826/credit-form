"""CLI entry point for form filler application."""
import sys
import shutil
from pathlib import Path
from openpyxl import load_workbook
from src.services.excel_processor import ExcelProcessor
from src.services.ai_generator import AIGenerator
from src.services.form_filler import FormFiller
from src.domain.exceptions import AIGenerationError, FieldDetectionError, ExcelProcessingError


def main(input_file: str, output_file: str = None):
    """Main CLI function."""
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)
    
    if output_file is None:
        output_file = input_path.stem + "_output" + input_path.suffix
    
    output_path = Path(output_file)
    
    try:
        print(f"Copying template file: {input_file} -> {output_file}")
        shutil.copy(input_file, output_file)
        
        print(f"Loading workbook: {output_file}")
        workbook = load_workbook(output_file)
        
        print("\n=== Detecting form fields ===")
        processor = ExcelProcessor(workbook)
        detection_result = processor.detect_fields()
        
        print(f"Total fields detected: {detection_result.count}")
        if detection_result.count == 0:
            print("No fields found in the worksheet. Exiting.")
            return
        
        print("\nDetected fields:")
        for label, field in detection_result.fields.items():
            print(f"  '{label}' -> {field.target_coordinate}")
        
        print("\n=== Generating data with AI ===")
        ai_generator = AIGenerator()
        field_labels = list(detection_result.fields.keys())
        data = ai_generator.generate_data(field_labels)
        
        print(f"Generated data for {len(data)} fields")
        
        print("\n=== Filling Excel with generated data ===")
        filler = FormFiller(workbook)
        fill_result = filler.fill(detection_result, data)
        
        print(f"\nSummary:")
        print(f"  Fields filled: {fill_result.filled_count}")
        print(f"  Fields skipped: {fill_result.skipped_count}")
        
        if fill_result.errors:
            print(f"\nErrors encountered:")
            for error in fill_result.errors:
                print(f"  - {error}")
        
        print(f"\nSaving Excel file: {output_file}")
        workbook.save(output_file)
        print(f"✓ Excel file saved as: {output_file}")
        
    except (AIGenerationError, FieldDetectionError, ExcelProcessingError) as e:
        print(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m src.cli.main <input_file> [output_file]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    main(input_file, output_file)

