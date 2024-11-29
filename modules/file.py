import os

def file(type, company):
    path = "assets/" + type + "_" + company + ".docx"
    if not os.path.exists(path):
        raise FileNotFoundError(f"Template file not found: {path}. Please ensure you have placed the template file in the assets directory.")
    return path