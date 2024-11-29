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
        "company_name": "pre",
        
        # Personal details
        "name": "Kirti Mishra",
        "employee_id": "PS100036/57",
        "designation": "US IT RECRUITER",
        "pan_no": "HMVPM3987G",
        "bank_no": "38943385226",
        
        # Salary details
        "initial_salary": 29300,
        "increment_amount": 350000,
        
        # Email details
        "email_id": "kirti.mishra@precisionstaffing.co.in",
        "phone_number": "9876543210",
        
        # Important dates
        "offer_date": "2022-01-01",
        "joining_date": "2022-01-10",
        "increment_letter_date": "2024-01-15",
        "increment_effective_date": "2024-02-01",
        "last_working_date": "2024-01-31",
        
        # Additional details
        "payslip_month": datetime.now().month, # Previous month
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
        email_date=employee["joining_date"],
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
        salary=employee["initial_salary"]
    )

    # Generate increment letter
    print("\n=== Generating Increment Letter ===")
    generate_increment_letter(
        company_name=employee["company_name"],
        letter_date=employee["increment_letter_date"],
        name=employee["name"],
        designation=employee["designation"],
        employee_id=employee["employee_id"],
        ctc=employee["increment_amount"],
        increment_date=employee["increment_effective_date"]
    )

    # Generate release letter
    print("\n=== Generating Release Letter ===")
    generate_release_letter(
        company_name=employee["company_name"],
        release_date=employee["last_working_date"],
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
