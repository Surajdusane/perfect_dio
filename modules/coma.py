import locale

def format_number_indian(number):
    # Set locale to India
    locale.setlocale(locale.LC_ALL, 'en_IN.UTF-8')
    
    # Format the number
    formatted_number = locale.format_string("%d", number, grouping=True)
    return formatted_number

def reverse_indian_format(formatted_number):
    # Remove any commas from the formatted number
    number_without_commas = formatted_number.replace(',', '')
    
    # Convert the string to integer
    original_number = int(number_without_commas)
    return original_number



# Example usage
# formatted_number = format_number_indian(1000)
# print(formatted_number)
