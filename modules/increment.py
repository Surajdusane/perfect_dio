from docxtpl import DocxTemplate
import os
from .coma import format_number_indian
from .file import file
from .date import convert_date

OUTPUT_DIR = "word"

def generate_increment_letter(
    company_name: str,
    letter_date: str,
    name: str,
    designation: str,
    employee_id: str,
    ctc: int,
    increment_date: str,
) -> None:
    """
    Generate an increment letter for an employee.
    
    Args:
        company_name: Name of the company
        letter_date: Date of the increment letter
        name: Employee name
        designation: Employee designation
        employee_id: Employee ID
        ctc: Cost to Company (annual salary)
        increment_date: Date when increment is effective
        year: Year of increment
    """
    try:
        # Ensure output directory exists
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Load and prepare document
        template_path = file("increment", company_name)
        doc = DocxTemplate(template_path)
        
        # Prepare context with formatted data
        context = {
            "dol": convert_date(letter_date),
            "name": name.title(),
            "empi": employee_id,
            "des": designation.upper(),
            "doi": convert_date(increment_date),
            "ctc": format_number_indian(ctc),
        }
        
        # Generate and save document
        doc.render(context)
        output_file = os.path.join(OUTPUT_DIR, f"Increment Letter - {name.title()}_{increment_date[:4]}.docx")
        doc.save(output_file)
        
        print(f"Success: Increment letter generated: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: Template 'increment_{company_name}.docx' not found in assets directory")
    except Exception as e:
        print(f"Error: {str(e)}")
