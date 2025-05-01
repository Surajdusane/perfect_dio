from modules.payslip import create_payslip
from modules.increment import generate_increment_letter
from modules.offer import generate_offer_letter
from modules.relese import generate_release_letter
from modules.email import generate_email
from modules.pdf_utils import convert_word_to_pdf, remove_pdf_metadata
from datetime import datetime, timedelta

def main():
    # Employee details with all dates and values
    employee = {
        # Company details
        "company_name": "asc",
        
        # Personal details
        "name": "Prashant B. Ratnaparkhi",
        "employee_id": "ASC100027/06",
        "designation": "IT RECRUITER",
        "pan_no": "CRTPR5413G",
        "bank_no": "00000032024295068",
        
        # Salary details
        "initial_salary": 19500,
        "payslip_ammount": 30500,
        "increment_amount": 24500,
        
        # Email details
        "email_id": "prashantratnaparkhi83@gmail.com",
        "phone_number": "7020652034",
        
        # Important dates
        "offer_date": "2022-09-26",
        "joining_date": "2022-10-03",
        "increment_letter_date": "2023-10-07",
        "increment_effective_date": "2023-10-07",
        "last_working_date": "2025-04-25",
        "release_date": "2025-04-25",
        "email_date": "2025-03-25",
        
        # Additional details
        "payslip_month": 5, 
        "x": "is"
        }
    
    # Generate offer letter
    print("\n=== Generating Offer Letter ===")
    generate_offer_letter(
        company_name=employee["company_name"],
        offer_date=employee["offer_date"],
        name=employee["name"],
        designation=employee["designation"],
        joining_date=employee["joining_date"],
        ctc=employee["initial_salary"] * 12
    )

    # Generate email credentials letter
    print("\n=== Generating Email Credentials Letter ===")
    generate_email(
        company_name=employee["company_name"],
        email_date=employee["email_date"],
        name=employee["name"],
        employee_id=employee["employee_id"],
        designation=employee["designation"],
        email_id=employee["email_id"],
        number=employee["phone_number"]
    )

    # Generate payslip
    print("\n=== Generating Payslips ===")
    create_payslip(
        company_name=employee["company_name"],
        month=employee["payslip_month"] - 1,
        date_of_joining=employee["joining_date"],
        pan_no=employee["pan_no"],
        bank_no=employee["bank_no"],
        name=employee["name"],
        employee_id=employee["employee_id"],
        designation=employee["designation"],
        salary=employee["payslip_ammount"]
    )

    # Generate increment letter
    print("\n=== Generating Increment Letter ===")
    generate_increment_letter(
        company_name=employee["company_name"],
        letter_date=employee["increment_letter_date"],
        name=employee["name"],
        designation=employee["designation"],
        employee_id=employee["employee_id"],
        ctc=employee["increment_amount"]*12,
        increment_date=employee["increment_effective_date"]
    )

    # Generate release letter
    print("\n=== Generating Release Letter ===")
    generate_release_letter(
        company_name=employee["company_name"],
        release_date=employee["release_date"],
        name=employee["name"],
        employee_id=employee["employee_id"],
        designation=employee["designation"],
        joining_date=employee["joining_date"],
        last_working_date=employee["last_working_date"],
        x=employee["x"]
    )

    # Convert all Word documents to PDF
    print("\n=== Converting Documents to PDF ===")
    convert_word_to_pdf()

    # Remove metadata from PDFs
    print("\n=== Removing PDF Metadata ===")
    remove_pdf_metadata()

if __name__ == "__main__":
    main()
