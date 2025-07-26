from docxtpl import DocxTemplate
import os
from .coma import format_number_indian
from .file import file
from .date import convert_date

OUTPUT_DIR = "word"

def generate_salary_certificatee(
    company_name: str,
    letter_date: str,
    name: str,
    designation: str,
    joining_date: str,
    current_ctc: int
) -> None:
    """
    Generate an Salary Certificate for an employee.
    
    Args:
        company_name: Name of the company
        offer_date: Date of the reveling
        name: Employee name
        designation: Employee designation
        joining_date: Date of joining
        ctc: Annual Cost to Company
    """
    try:
        # Ensure output directory exists
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Load template
        template_path = file("salary_certificate", company_name)
        doc = DocxTemplate(template_path)
        
        # Calculate salary components
        monthly_salary = current_ctc / 12
        basic_pay = (monthly_salary - 2500) * 0.70
        hra = (monthly_salary - 2500) * 0.18
        bonus = (monthly_salary - 2500) * 0.12
        special_allowance = 2500
        net_monthly = monthly_salary - 800
        
        # Prepare context with formatted data
        context = {
            "dol": convert_date(letter_date),
            "name": name.title(),
            "des": designation.upper(),
            "doj": convert_date(joining_date),
            "ctc": format_number_indian(current_ctc),
            "mbs": int(basic_pay),
            "abs": int(basic_pay * 12),
            "mhra": int(hra),
            "ahra": int(hra * 12),
            "msb": int(bonus),
            "asb": int(bonus * 12),
            "msa": int(special_allowance),
            "asa": int(special_allowance * 12),
            "mgs": int(monthly_salary),
            "ags": current_ctc,
            "mctc": int(net_monthly),
            "actc": int(net_monthly * 12),
        }
        
        # Generate and save document
        doc.render(context)
        output_file = os.path.join(OUTPUT_DIR, f"Salary Certificate - {name.title()}.docx")
        doc.save(output_file)
        
        print(f"Success: Offer letter generated: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: Template 'salary_certificate_{company_name}.docx' not found in assets directory")
    except Exception as e:
        print(f"Error: {str(e)}")
