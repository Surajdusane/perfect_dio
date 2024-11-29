from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from datetime import datetime
import os

from modules.payslip import create_payslip
from modules.increment import generate_increment_letter
from modules.offer import generate_offer_letter
from modules.relese import generate_release_letter
from modules.email import generate_email
from modules.pdf_utils import convert_word_to_pdf, remove_pdf_metadata
from modules.file_utils import clear_directories, create_zip_archive

app = FastAPI(
    title="Employee Document Generator API",
    description="API for generating various employee documents including offer letters, payslips, and more",
    version="1.0.0"
)

class EmployeeData(BaseModel):
    # Company details
    company_name: str
    
    # Personal details
    name: str
    employee_id: str
    designation: str
    pan_no: str
    bank_no: str
    
    # Salary details
    initial_salary: float
    increment_amount: float
    
    # Email details
    email_id: str
    number: str
    
    # Important dates
    offer_date: str
    joining_date: str
    increment_letter_date: str
    increment_effective_date: str
    last_working_date: str
    
    # Additional details
    payslip_month: Optional[int] = None
    x: str

    class Config:
        json_schema_extra = {
            "example": {
                "company_name": "pre",
                "name": "Kirti Mishra",
                "employee_id": "PS100036/57",
                "designation": "US IT RECRUITER",
                "pan_no": "HMVPM3987G",
                "bank_no": "38943385226",
                "initial_salary": 29300,
                "increment_amount": 35000,
                "email_id": "kirti.mishra@precisionstaffing.co.in",
                "number": "9579330721",
                "offer_date": "2022-01-01",
                "joining_date": "2022-01-10",
                "increment_letter_date": "2024-01-15",
                "increment_effective_date": "2024-02-01",
                "last_working_date": "2024-01-31",
                "payslip_month": 1,
                "x": "his"
            }
        }

@app.post("/generate-documents")
async def generate_documents(employee: EmployeeData):
    try:
        # Set default payslip month if not provided
        if employee.payslip_month is None:
            employee.payslip_month = datetime.now().month

        # Generate offer letter
        print("\n=== Generating Offer Letter ===")
        generate_offer_letter(
            company_name=employee.company_name,
            offer_date=employee.offer_date,
            name=employee.name,
            designation=employee.designation,
            joining_date=employee.joining_date,
            ctc=employee.initial_salary * 12
        )

        # Generate email credentials letter
        print("\n=== Generating Email Credentials Letter ===")
        generate_email(
            company_name=employee.company_name,
            email_date=employee.joining_date,
            name=employee.name,
            employee_id=employee.employee_id,
            designation=employee.designation,
            email_id=employee.email_id,
            number=employee.number
        )

        # Generate payslip
        print("\n=== Generating Payslips ===")
        create_payslip(
            company_name=employee.company_name,
            month=employee.payslip_month - 1,
            date_of_joining=employee.joining_date,
            pan_no=employee.pan_no,
            bank_no=employee.bank_no,
            name=employee.name,
            employee_id=employee.employee_id,
            designation=employee.designation,
            salary=employee.initial_salary
        )

        # Generate increment letter
        print("\n=== Generating Increment Letter ===")
        generate_increment_letter(
            company_name=employee.company_name,
            letter_date=employee.increment_letter_date,
            name=employee.name,
            designation=employee.designation,
            employee_id=employee.employee_id,
            ctc=employee.increment_amount,
            increment_date=employee.increment_effective_date
        )

        # Generate release letter
        print("\n=== Generating Release Letter ===")
        generate_release_letter(
            company_name=employee.company_name,
            release_date=employee.last_working_date,
            name=employee.name,
            employee_id=employee.employee_id,
            designation=employee.designation,
            joining_date=employee.joining_date,
            last_working_date=employee.last_working_date,
            x=employee.x
        )

        # Convert all Word documents to PDF
        print("\n=== Converting Documents to PDF ===")
        convert_word_to_pdf()

        # Remove metadata from PDFs
        print("\n=== Removing PDF Metadata ===")
        remove_pdf_metadata()

        return {
            "status": "success",
            "message": "All documents generated successfully",
            "documents": [
                "Offer Letter",
                "Email Credentials Letter",
                "Payslips (6 months)",
                "Increment Letter",
                "Release Letter"
            ]
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/clear-files")
async def clear_files():
    """Clear all generated files from word and pdf directories."""
    try:
        result = clear_directories()
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download-documents")
async def download_documents(background_tasks: BackgroundTasks):
    """Download all generated PDF documents as a ZIP file."""
    try:
        # Create ZIP archive
        zip_path = create_zip_archive()
        
        # Return ZIP file
        if os.path.exists(zip_path):
            def cleanup_zip():
                try:
                    if os.path.exists(zip_path):
                        os.unlink(zip_path)
                except Exception as e:
                    print(f"Error cleaning up ZIP file: {e}")

            background_tasks.add_task(cleanup_zip)
            
            return FileResponse(
                path=zip_path,
                media_type="application/zip",
                filename="documents.zip"
            )
        else:
            raise HTTPException(status_code=404, detail="No documents found")
            
    except Exception as e:
        if os.path.exists(zip_path):
            os.unlink(zip_path)
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, port=8000)
