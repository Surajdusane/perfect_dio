from docxtpl import DocxTemplate
import os
from .coma import format_number_indian
from .file import file
from .date import convert_date

OUTPUT_DIR = "word"

def generate_experience_letter(
    company_name: str,
    release_date: str,
    name: str,
    employee_id: str,
    designation: str,
    joining_date: str,
    last_working_date: str,
    x: str
) -> None:
    """
    Generate a experience letter for an employee.
    
    Args:
        company_name: Name of the company
        experience_date: Date of the experience letter
        name: Employee name
        employee_id: Employee ID
        designation: Employee designation
        joining_date: Date of joining
        last_working_date: Last working date
    """
    try:
        # Ensure output directory exists
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Load template
        template_path = file("experience", company_name)
        doc = DocxTemplate(template_path)
        
        # Prepare context with formatted data
        context = {
            "date": convert_date(release_date),
            "name": name.title(),
            "emi": employee_id,
            "des": designation.upper(),
            "jdate": convert_date(joining_date),
            "rdate": convert_date(last_working_date),
            "x": x
        }
        
        # Generate and save document
        doc.render(context)
        output_file = os.path.join(OUTPUT_DIR, f"Experience Letter - {name.title()}.docx")
        doc.save(output_file)
        
        print(f"Success: experience letter generated: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: Template 'Experience Letter of {company_name}.docx' not found in assets directory")
    except Exception as e:
        print(f"Error: {str(e)}")
