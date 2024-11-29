from datetime import datetime

def convert_date(date_str):
    return datetime.strptime(date_str, '%Y-%m-%d').strftime('%B %d, %Y')
