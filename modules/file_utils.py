import os
import shutil
from pathlib import Path
import zipfile

def clear_directories():
    """Clear all files from word and pdf directories."""
    directories = ['word', 'pdf']
    for dir_name in directories:
        if os.path.exists(dir_name):
            for file in os.listdir(dir_name):
                file_path = os.path.join(dir_name, file)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Error: {e}")
    return {"message": "All files cleared successfully"}

def create_zip_archive():
    """Create a ZIP archive of all PDF files."""
    pdf_dir = "pdf"
    zip_path = "documents.zip"
    
    # Create a ZIP file
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add all PDF files
        if os.path.exists(pdf_dir):
            for file in os.listdir(pdf_dir):
                if file.endswith('.pdf'):
                    file_path = os.path.join(pdf_dir, file)
                    zipf.write(file_path, os.path.basename(file_path))
    
    return zip_path
