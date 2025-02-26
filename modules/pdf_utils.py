import os
import win32com.client
import pikepdf
from pathlib import Path

WORD_DIR = "word"
PDF_DIR = "pdf"

def convert_word_to_pdf():
    """
    Convert all Word documents in the word directory to PDF format.
    Saves the PDFs in a 'pdf' directory.
    """
    try:
        # Ensure PDF directory exists
        os.makedirs(PDF_DIR, exist_ok=True)
        
        # Initialize Word application
        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        
        # Convert each Word document
        for file in os.listdir(WORD_DIR):
            if file.endswith(".docx"):
                # Full paths
                docx_path = os.path.abspath(os.path.join(WORD_DIR, file))
                pdf_path = os.path.abspath(os.path.join(PDF_DIR, file.replace(".docx", ".pdf")))
                
                try:
                    # Load and convert
                    doc = word.Documents.Open(docx_path)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
                    doc.Close()
                    print(f"Success: Converted {file} to PDF")
                except Exception as e:
                    print(f"Error converting {file}: {str(e)}")
        
        # Cleanup
        word.Quit()
        print("\nAll Word documents converted to PDF")
        
    except Exception as e:
        print(f"Error in conversion process: {str(e)}")

def remove_pdf_metadata():
    """
    Remove metadata from all PDF files in the pdf directory.
    Creates clean copies of the PDFs with no metadata.
    """
    try:
        # Process each PDF file
        for file in os.listdir(PDF_DIR):
            if file.endswith(".pdf"):
                pdf_path = os.path.join(PDF_DIR, file)
                temp_path = os.path.join(PDF_DIR, f"temp_{file}")
                
                try:
                    # Open and clean the PDF
                    with pikepdf.open(pdf_path) as pdf:
                        # Create new PDF with minimal metadata
                        with pikepdf.new() as new_pdf:
                            new_pdf.pages.extend(pdf.pages)
                            # Save to temporary file
                            new_pdf.save(temp_path)
                    
                    # Replace original with clean version
                    os.replace(temp_path, pdf_path)
                    print(f"Success: Removed metadata from {file}")
                    
                except Exception as e:
                    print(f"Error processing {file}: {str(e)}")
                    # Clean up temp file if it exists
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
        
        print("\nMetadata removed from all PDF files")
        
    except Exception as e:
        print(f"Error in metadata removal process: {str(e)}")

# convert_word_to_pdf()
# remove_pdf_metadata()