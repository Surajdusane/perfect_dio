from docxtpl import DocxTemplate
import os
folder_path = "word"
from .date import convert_date
from .coma import format_number_indian
from .file import file

month_details = {
    1: ["January", 31],
    2: ["February", 29],
    3: ["March", 31],
    4: ["April", 30],
    5: ["May", 31],
    6: ["June", 30],
    7: ["July", 31],
    8: ["August", 31],
    9: ["September", 30],
    10: ["October", 31],
    11: ["November", 30],
    12: ["December", 31]
}


def calculate_PT(k3):
    """Returns the PT based on the value of k3."""
    if k3 <= 10000:
        return 0
    elif k3 <= 15000:
        return 150
    else:
        return 200

def get_next_month_name(current_month):
    # Dictionary of month numbers to names
    months = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December"
    }
    
    # Calculate next month number (wrapping around to 1 if current month is 12)
    next_month = current_month + 1 if current_month < 12 else 1
    
    # Return the name of the next month
    return months[next_month]

def last6monthdata(month):
    """Returns the pay period and days for the last 6 months based on the input month."""
    data = {}
    
    # Loop through the last 6 months, handling the month wrapping using modulo
    for i in range(6):
        current_month = ((month - i - 1) % 12) + 1  # Convert 0 to 12 for December
        data[current_month] = month_details[current_month][1]
    
    return data


def create_payslip(company_name, month, date_of_joining, pan_no, bank_no, name, employee_id, designation, salary, year=2024):
    try:
        template_path = file("payslip", company_name)
        doc = DocxTemplate(template_path)
        
        basic_pay = (int(salary)-2500)*70/100
        hra = (int(salary)-2500)*18/100
        bonus = (int(salary)-2500)*12/100
        total_earning = salary #int(basic_pay) + int(hra) + int(bonus) + 2500
        pt = calculate_PT(int(salary))
        total_deduction = 800
        net_pay = int(salary) - total_deduction

        # Create the word directory if it doesn't exist
        os.makedirs(folder_path, exist_ok=True)

        # Generate payslips for last 6 months
        for i in range(6):
            # Calculate current month for this iteration
            current_month = ((month - i) % 12) or 12  # Convert 0 to 12 for December
            
            # Calculate previous month
            previous_month = ((current_month - 1) % 12) or 12  # Convert 0 to 12 for December
            
            context = {
                'mon': month_details[current_month][0],  # Current month
                'date': convert_date(date_of_joining),
                'year': year if current_month > 5 else year + 1,
                'pep': month_details[previous_month][0],  # Previous month
                'pd': month_details[previous_month][1],  # Days in previous month
                'pan': pan_no,
                'bank': bank_no,
                'name': name.upper(),
                'empi': employee_id,
                'des': designation,
                'bp': format_number_indian(int(basic_pay)),
                'hra': format_number_indian(int(hra)),
                'bonus': format_number_indian(int(bonus)),
                'te': format_number_indian(int(total_earning)),
                'pt': pt,
                'td': total_deduction,
                'np': format_number_indian(int(net_pay)),
                "loc": "Kolkata",
            }
            
            # Create a new template for each month to avoid overwriting
            doc = DocxTemplate(template_path)
            doc.render(context)
            
            # Save the generated payslip with current month in filename
            file_path = os.path.join(folder_path, f"Payslip of {month_details[current_month][0]}.docx")
            doc.save(file_path)
            print(f"Generated payslip for {month_details[current_month][0]} (Pay period: {month_details[previous_month][0]})")
            
    except FileNotFoundError as e:
        print(f"Error: {str(e)}")
        print(f"Please ensure you have placed the template file 'payslip_{company_name}.docx' in the assets directory.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
