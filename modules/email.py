from docxtpl import DocxTemplate
import os
from .coma import format_number_indian
from .file import file
from .date import convert_date
from datetime import datetime

OUTPUT_DIR = "word"

def generate_email(
    company_name: str,
    email_date: str,
    name: str,
    employee_id: str,
    designation: str,
    email_id: str,
    number: str
) -> None:
    """
    Generate an email credentials letter for an employee.
    
    Args:
        company_name: Name of the company
        email_date: Date of the email letter
        name: Employee name
        employee_id: Employee ID
        designation: Employee designation
        email_id: Company email ID
        password: Initial password
    """
    try:
        # Ensure output directory exists
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Load template
        template_path = file("email", company_name)
        doc = DocxTemplate(template_path)
        
        # Prepare context with formatted data
        context = {
            "name": name.title(),
            "mail": email_id,
            "outdate": datetime.strptime(email_date, '%Y-%m-%d').strftime('%d/%m/%Y'),
            "dts": convert_date(email_date),
            "dtr": convert_date(email_date),
            "empi": employee_id,
            "des": designation.upper(),
            "num": number
        }
        
        # Generate and save document
        doc.render(context)
        output_file = os.path.join(OUTPUT_DIR, f"Mail - {name.title()}.docx")
        doc.save(output_file)
        
        print(f"Success: Email credentials letter generated: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: Template 'Resignation Letter of {company_name}.docx' not found in assets directory")
    except Exception as e:
        print(f"Error: {str(e)}")
