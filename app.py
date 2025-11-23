from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import range_boundaries, get_column_letter
import json
import os
import shutil
from openai import OpenAI

# Copy the template file first to preserve all formatting
shutil.copy("sample.xlsx", "sample_output.xlsx")

# Now open the copy and work with it
wb = load_workbook("sample_output.xlsx")
ws = wb.active

merged_map = {}

for r in ws.merged_cells.ranges:
    merged_map[r.coord] = r.coord.split(":")[0]  # top-left cell

print("Merged cells found:")
print(f"Total merged ranges: {len(merged_map)}")
for merged_range, top_left in merged_map.items():
    print(f"  {merged_range} -> {top_left}")

if not merged_map:
    print("No merged cells found in the worksheet.")


# Spanish form field keywords commonly found in bank documents
LABEL_KEYWORDS = [
    'nombre', 'dirección', 'teléfono', 'telefono', 'email', 'correo', 'fecha',
    'código', 'codigo', 'actividad', 'representante', 'razón', 'razon', 'social',
    'nit', 'rfc', 'curp', 'rfc', 'ciudad', 'estado', 'país', 'pais', 'cp', 'código postal',
    'banco', 'cuenta', 'clabe', 'swift', 'iban', 'moneda', 'monto', 'importe',
    'apellido', 'paterno', 'materno', 'nacimiento', 'edad', 'género', 'genero',
    'ocupación', 'ocupacion', 'profesión', 'profesion', 'empresa', 'puesto',
    'documento', 'identificación', 'identificacion', 'pasaporte', 'licencia',
    'contacto', 'emergencia', 'parentesco', 'beneficiario', 'titular',
    'firma', 'fecha de', 'lugar de', 'hora', 'folio', 'referencia', 'número', 'numero'
]

# Collect all potential fields first (before filtering)
all_potential_fields = {}

for row in ws.iter_rows():
    for cell in row:
        if cell.value and isinstance(cell.value, str):
            label = cell.value.strip()
            
            # find the field cell(s)
            row_idx = cell.row
            col_idx = cell.column
            
            # field is usually the next cell to the right
            field_cell = ws.cell(row=row_idx, column=col_idx+1)
            
            # detect merged region
            target = None
            for m in merged_map:
                top_left = merged_map[m]
                if ws[top_left].row == row_idx and ws[top_left].column == col_idx+1:
                    target = m  # merged range
                    break
            
            if target is None:
                target = field_cell.coordinate  # single cell
            
            all_potential_fields[label] = target

print("\n=== FIELD DETECTION ANALYSIS ===")
print(f"Total potential fields found (before filtering): {len(all_potential_fields)}")

# Now apply filtering logic
fields = {}

for label, target in all_potential_fields.items():
    # Filter condition 1: Skip cells with less than 3 characters
    if len(label) < 3:
        continue
    
    # Filter condition 2: Only include cells with form field keywords or ending with ':'
    label_lower = label.lower().strip()
    has_keyword = any(keyword in label_lower for keyword in LABEL_KEYWORDS)
    ends_with_colon = label.rstrip().endswith(':')
    
    if not (has_keyword or ends_with_colon):
        continue
    
    # Filter condition 3: Check if the next cell to the right is EMPTY
    # If it already has data, skip it since it's not a fillable field
    try:
        # Extract row and column from target
        if ":" in target:
            # For merged cells, get the top-left cell
            top_left = target.split(":")[0]
            field_cell = ws[top_left]
        else:
            field_cell = ws[target]
        
        # Check if the field cell is empty (or contains only whitespace)
        field_value = field_cell.value
        is_empty = field_value is None or (isinstance(field_value, str) and field_value.strip() == "")
        
        if not is_empty:
            # Field already has data, skip it
            continue
    except Exception as e:
        # If we can't check the cell, skip it to be safe
        print(f"  Warning: Could not check field cell for '{label}': {e}")
        continue
    
    # All conditions passed - add to fields dictionary
    fields[label] = target

print(f"Total fields found (after filtering): {len(fields)}")
print(f"Fields filtered out: {len(all_potential_fields) - len(fields)}")
if len(all_potential_fields) > 0:
    print(f"Filtering efficiency: {((len(all_potential_fields) - len(fields)) / len(all_potential_fields) * 100):.1f}% reduction")
else:
    print("Filtering efficiency: N/A (no potential fields found)")

print("\nFiltered fields extracted:")
for label, target in fields.items():
    print(f"  '{label}' -> {target}")

if not fields:
    print("No fields found in the worksheet after filtering.")


def get_data_from_ai(field_labels, max_retries=3):
    """
    Calls AI API to generate JSON data based on the field labels.
    Returns a dictionary with field labels as keys and generated values.
    """
    # Initialize OpenAI client (requires OPENAI_API_KEY environment variable)
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set. Please set it to use AI data generation.")
    
    client = OpenAI(api_key=api_key)
    
    # Create prompt for AI - be more explicit about JSON format
    field_list = ", ".join(field_labels)
    prompt = f"""Generate realistic sample data for the following form fields: {field_list}

IMPORTANT: Return ONLY a complete, valid JSON object. The JSON must:
1. Have each field name as a key (exactly as listed above)
2. Have appropriate sample data as the value for each key
3. Be complete and properly closed (all brackets and braces must be closed)
4. Be valid JSON that can be parsed

Return ONLY the JSON object, no explanations, no markdown, no code blocks.

Example format:
{{"Full Name": "John Doe", "Address": "123 Main St, City, Country", "DOB": "01-15-1990", "Gender": "Male"}}
"""
    
    print("\nCalling AI to generate data...")
    print(f"Fields sent to AI: {field_list}")
    print(f"Total fields: {len(field_labels)}")
    
    for attempt in range(max_retries):
        try:
            # Calculate appropriate max_tokens based on number of fields
            # Estimate: ~50-100 tokens per field + buffer for structure
            estimated_tokens = max(2000, len(field_labels) * 80 + 1000)
            print(f"Using max_tokens: {estimated_tokens} (attempt {attempt + 1}/{max_retries})")
            
            # Try with JSON mode first, fallback if not supported
            try:
                response = client.chat.completions.create(
                    model="gpt-4.1-mini",
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant that generates realistic sample data. Always return complete, valid JSON objects. Never return incomplete or truncated JSON. Ensure all field names from the user's list are included as keys."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.2,
                    max_tokens=estimated_tokens,
                    response_format={"type": "json_object"}  # Force JSON mode
                )
            except Exception as json_mode_error:
                # If JSON mode not supported, try without it
                print(f"JSON mode not supported, trying without it: {json_mode_error}")
                response = client.chat.completions.create(
                    model="gpt-4.1-mini",
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant that generates realistic sample data. Always return complete, valid JSON objects. Never return incomplete or truncated JSON. Ensure all field names from the user's list are included as keys."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.2,
                    max_tokens=estimated_tokens
                )
            
            # Extract JSON from response
            ai_response = response.choices[0].message.content.strip()
            
            # Remove markdown code blocks if present
            if ai_response.startswith("```"):
                parts = ai_response.split("```")
                # Find the JSON part (usually the middle part)
                for part in parts:
                    part = part.strip()
                    if part.startswith("json"):
                        part = part[4:].strip()
                    if part.startswith("{") or part.startswith("["):
                        ai_response = part
                        break
                else:
                    # If no JSON found, try the first non-empty part
                    ai_response = next((p.strip() for p in parts if p.strip() and not p.strip().startswith("json")), ai_response)
            
            # Try to fix incomplete JSON if needed
            ai_response = ai_response.strip()
            if not ai_response.endswith("}") and not ai_response.endswith("]"):
                # Try to find and close incomplete JSON
                open_braces = ai_response.count("{") - ai_response.count("}")
                open_brackets = ai_response.count("[") - ai_response.count("]")
                
                if open_braces > 0 or open_brackets > 0:
                    print(f"Warning: JSON appears incomplete. Attempting to fix...")
                    # Close any open structures
                    ai_response += "}" * open_braces + "]" * open_brackets
            
            # Parse JSON
            data = json.loads(ai_response)
            
            print(f"AI Response received successfully: {len(data)} fields")
            print(f"AI Response preview (first 3 fields): {json.dumps(dict(list(data.items())[:3]), indent=2)}...")
            return data
            
        except json.JSONDecodeError as e:
            print(f"Error parsing AI response as JSON (attempt {attempt + 1}/{max_retries}): {e}")
            if attempt < max_retries - 1:
                print(f"Retrying with increased token limit...")
                # Increase tokens on retry
                estimated_tokens = int(estimated_tokens * 1.5)
                continue
            else:
                print(f"Raw AI response (last 500 chars): {ai_response[-500:]}")
                raise ValueError(f"Failed to parse AI response as JSON after {max_retries} attempts. The response may be incomplete. Last error: {e}")
        except Exception as e:
            print(f"Error calling AI API (attempt {attempt + 1}/{max_retries}): {e}")
            if attempt < max_retries - 1:
                print(f"Retrying...")
                continue
            raise


# Get field labels (keys from fields dictionary)
field_labels = list(fields.keys())

if field_labels:
    # Call AI to get dynamic data
    data = get_data_from_ai(field_labels)
    
    print("\nFilling Excel with AI-generated data...")
    
    # Helper function to check if a cell coordinate is within a merged range
    def is_cell_in_range(cell_coord, merged_range):
        """Check if a cell coordinate is within a merged range"""
        if ":" not in merged_range:
            return cell_coord == merged_range
        
        start, end = merged_range.split(":")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(merged_range)
            cell_col = ws[cell_coord].column
            cell_row = ws[cell_coord].row
            return min_col <= cell_col <= max_col and min_row <= cell_row <= max_row
        except:
            # Fallback: simple string check
            return cell_coord == start or (start <= cell_coord <= end)
    
    # Fill Excel with the AI-generated data
    filled_count = 0
    skipped_count = 0
    
    for label, target in fields.items():
        try:
            if label in data:
                top_left = target.split(":")[0]  # merged or single
                value = data[label]
                
                # Convert complex objects (dict, list) to JSON string
                if isinstance(value, (dict, list)):
                    value = json.dumps(value, ensure_ascii=False)
                # Convert None to empty string
                elif value is None:
                    value = ""
                
                # Get the cell - handle merged cells properly
                # If target is a merged range, use the top-left from merged_map
                if ":" in target and target in merged_map:
                    # This is a merged range, use the stored top-left cell
                    writable_cell_coord = merged_map[target]
                else:
                    # Single cell or not in merged_map, use the coordinate directly
                    writable_cell_coord = top_left
                
                cell = ws[writable_cell_coord]
                
                # If it's a MergedCell, find the actual top-left cell
                if isinstance(cell, MergedCell):
                    found_top_left = None
                    # Find which merged range contains this cell
                    for merged_range, merged_top_left in merged_map.items():
                        if is_cell_in_range(writable_cell_coord, merged_range):
                            found_top_left = merged_top_left
                            break
                    
                    if found_top_left:
                        cell = ws[found_top_left]
                    else:
                        # If we can't find the top-left, try to get it from the merged_cells
                        for merged_range_obj in ws.merged_cells.ranges:
                            # Check if our coordinate is in this range
                            try:
                                range_str = f"{get_column_letter(merged_range_obj.min_col)}{merged_range_obj.min_row}:{get_column_letter(merged_range_obj.max_col)}{merged_range_obj.max_row}"
                                if is_cell_in_range(writable_cell_coord, range_str):
                                    # Get the top-left cell of this merged range
                                    cell = ws.cell(row=merged_range_obj.min_row, column=merged_range_obj.min_col)
                                    break
                            except:
                                continue
                        else:
                            # Skip this cell if we can't find a writable cell
                            print(f"  Skipped '{label}' -> {writable_cell_coord} (cannot find writable cell)")
                            skipped_count += 1
                            continue
                
                # Set the value on the writable cell
                cell.value = value
                filled_count += 1
                print(f"  Filled '{label}' -> {cell.coordinate} with: {str(value)[:50]}...")
            else:
                print(f"  Warning: '{label}' not found in AI response data")
        except Exception as e:
            print(f"  Error filling '{label}': {e}")
            skipped_count += 1
            continue  # Continue with next field even if this one fails
    
    print(f"\nSummary: {filled_count} fields filled, {skipped_count} fields skipped")
    
    # Save the output file
    output_filename = "sample_output.xlsx"
    wb.save(output_filename)
    print(f"\nExcel file saved as: {output_filename}")
else:
    print("\nNo fields to process. Skipping AI data generation.")