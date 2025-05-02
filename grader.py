import os
import io
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx2pdf import convert
from docx import Document

# >>> Using your new library approach:
from openai import OpenAI

client = OpenAI(api_key="")

def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF, including OCR on any embedded images using pytesseract.
    The result is a single combined string containing page text + any OCR text.
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"[Error opening PDF file {pdf_path}: {e}]")
        return ""
    
    all_text = ""
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        # Extract normal text
        page_text = page.get_text()
        all_text += page_text

        # Extract images from the page and do OCR
        image_list = page.get_images(full=True)
        if image_list:
            all_text += f"\n[OCR from images on page {page_idx}]:\n"
        for img in image_list:
            xref = img[0]
            try:
                base_img = doc.extract_image(xref)
                img_bytes = base_img["image"]
                pil_img = Image.open(io.BytesIO(img_bytes))
                ocr_text = pytesseract.image_to_string(pil_img)
                all_text += ocr_text + "\n"
            except Exception as ex:
                all_text += f"[Error processing image: {ex}]\n"
    return all_text

def extract_text_from_docx(docx_path):
    """
    Extract text from .doc / .docx using python-docx (images in docx aren't OCR'd).
    """
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"[Error opening DOCX file {docx_path}: {e}]")
        return ""
    
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n".join(paragraphs)

def convert_word_to_pdf(word_path, out_folder):
    """
    Convert .doc/.docx -> PDF using docx2pdf. Return the PDF path or None if it fails.
    """
    try:
        base_name = os.path.splitext(os.path.basename(word_path))[0]
        out_pdf = os.path.join(out_folder, f"{base_name}.pdf")
        convert(word_path, out_pdf)
        print(f"Converted '{word_path}' to PDF: '{out_pdf}'")
        return out_pdf
    except Exception as e:
        print(f"[Error converting '{word_path}' to PDF: {e}]")
        return None

def parse_lab_report(file_path, converted_folder):
    """
    1. If it's already PDF, parse it directly with PyMuPDF + Tesseract.
    2. If it's doc/docx, attempt docx2pdf -> parse the result as PDF.
    3. If conversion fails, parse docx text directly (no images).
    Returns the final text (typed + OCR).
    """
    f_lower = file_path.lower()
    if f_lower.endswith(".pdf"):
        return extract_text_from_pdf(file_path)
    else:
        # doc/docx => try converting to PDF
        pdf_candidate = convert_word_to_pdf(file_path, converted_folder)
        if pdf_candidate:
            return extract_text_from_pdf(pdf_candidate)
        else:
            # fallback => docx text only
            return extract_text_from_docx(file_path)

def grade_lab_report(rubric_text, lab_manual_text, student_text):
    """
    Calls client.responses.create(...) with an array of messages.
    We pass the rubric, lab manual, and student text together,
    and ask for a grade out of 30 plus feedback.
    """
    # We'll build a single user message that includes everything:
    user_content = f"""
Please grade this lab report.

Rubric:
{rubric_text}

Lab Manual:
{lab_manual_text}

Student Lab Report (including OCR'd images):
{student_text}

Please provide:
1) A final grade out of 30.
2) Detailed feedback on strengths/weaknesses.
"""

    messages_input = [
        {
            "role": "system",
            "content": "You are a meticulous and precise grading assistant. Please be consistent on the grading for each student"
        },
        {
            "role": "user",
            "content": user_content
        }
    ]

    response = client.responses.create(
        model="gpt-4.1",   # or whichever model is valid in your environment
        input=messages_input
    )

    return response.output_text

def main(folder_path):
    """
    1) Locate 'Rubrics.docx' and 'LabX.pdf' for the official rubric & manual.
    2) For each other doc/pdf, parse the text (inc. OCR).
    3) Call grade_lab_report to get a final grade + feedback; 
    4) Save the grading result to the 'graded_reports' subfolder.
    """
    rubric_text = None
    lab_manual_text = None

    # Where to place converted PDFs from docx
    converted_dir = os.path.join(folder_path, "converted_pdf")
    os.makedirs(converted_dir, exist_ok=True)

    # Where to save grading results
    graded_dir = os.path.join(folder_path, "graded_reports")
    os.makedirs(graded_dir, exist_ok=True)

    # Step 1: Find Rubrics.docx and LabX.pdf
    for fn in os.listdir(folder_path):
        f_lower = fn.lower()
        full_path = os.path.join(folder_path, fn)
        if "rubrics" in f_lower and f_lower.endswith(".docx"):
            rubric_text = extract_text_from_docx(full_path)
        elif f_lower.startswith("lab") and f_lower.endswith(".pdf"):
            lab_manual_text = extract_text_from_pdf(full_path)

    if not rubric_text:
        print("Error: Could not find 'Rubrics.docx' in the folder.")
        return
    if not lab_manual_text:
        print("Error: Could not find a 'LabX.pdf' (lab manual).")
        return

    # Step 2: Parse each student's doc/pdf, ignoring the known rubric/manual
    for fn in os.listdir(folder_path):
        f_lower = fn.lower()
        if "rubrics" in f_lower and f_lower.endswith(".docx"):
            continue
        if f_lower.startswith("lab") and f_lower.endswith(".pdf"):
            continue
        # skip anything not doc/pdf
        if not (f_lower.endswith(".pdf") or f_lower.endswith(".doc") or f_lower.endswith(".docx")):
            continue

        full_path = os.path.join(folder_path, fn)
        student_text = parse_lab_report(full_path, converted_dir)

        # If the student submission yields no text, skip
        if not student_text.strip():
            print(f"[Skipping '{fn}' because no text was extracted.]")
            continue

        # Step 3: Grade it
        grading_result = grade_lab_report(rubric_text, lab_manual_text, student_text)

        # Step 4: Save to 'graded_reports'
        base_name = os.path.splitext(fn)[0]
        out_file = os.path.join(graded_dir, f"{base_name}_grade.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(grading_result)

        print(f"\n=== FINAL GRADE FOR {fn} ===\n")
        print(grading_result)
        print(f"\n[Result saved to: {out_file}]\n")
        print("\n================================\n")


if __name__ == "__main__":
    # Example usage:
    folder_path = "Lab08/"
    main(folder_path)
