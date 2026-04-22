import os
import sys

def try_extract(pdf_path):
    try:
        import fitz
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text() + "\n"
        print("--- START TEXT ---")
        print(text)
        print("--- END TEXT ---")
        return
    except ImportError:
        pass
    except Exception as e:
        print(f"fitz error: {e}")

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            print("--- START TEXT ---")
            print(text)
            print("--- END TEXT ---")
            return
    except ImportError:
        print("Neither fitz nor pdfplumber found")
    except Exception as e:
        print(f"pdfplumber error: {e}")

if __name__ == "__main__":
    path = r"C:\Users\sdighero\ZenflowProjects\vorrei-migliorare-il-funzionamen-1b43\.zenflow-attachments\f12ca45b-9428-41cd-a645-5c0d17f94227.pdf"
    try_extract(path)
