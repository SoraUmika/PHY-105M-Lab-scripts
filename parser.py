import os
import io
import fitz  # PyMuPDF for PDF processing
import pytesseract
from PIL import Image
from docx2pdf import convert
from docx import Document

def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF file {pdf_path}: {e}")
        return ""
    
    all_text = ""
    for page_index in range(len(doc)):
        page = doc[page_index]
        page_text = page.get_text()
        all_text += page_text

        # OCR any images on this page
        image_list = page.get_images(full=True)
        if image_list:
            all_text += "\n[OCR extracted from images on this page:]\n"
        for img in image_list:
            xref = img[0]
            try:
                base_img = doc.extract_image(xref)
                img_bytes = base_img["image"]
                pil_img = Image.open(io.BytesIO(img_bytes))
                ocr_text = pytesseract.image_to_string(pil_img)
                all_text += ocr_text + "\n"
            except Exception as ex:
                all_text += f"\n[Error processing image: {ex}]\n"
    return all_text

def extract_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"Error opening DOCX file {docx_path}: {e}")
        return ""
    
    paragraphs = [para.text for para in doc.paragraphs]
    return "\n".join(paragraphs)

def convert_word_to_pdf(word_path, out_folder):
    try:
        base = os.path.splitext(os.path.basename(word_path))[0]
        output_pdf = os.path.join(out_folder, f"{base}.pdf")
        convert(word_path, output_pdf)
        print(f"Converted '{word_path}' to PDF: '{output_pdf}'")
        return output_pdf
    except Exception as e:
        print(f"Error converting '{word_path}' to PDF: {e}")
        return None

def main(folder_path):
    converted_dir = os.path.join(folder_path, "converted_pdf")
    parsed_dir = os.path.join(folder_path, "parsed_docs")

    os.makedirs(converted_dir, exist_ok=True)
    os.makedirs(parsed_dir, exist_ok=True)

    for filename in os.listdir(folder_path):
        f_lower = filename.lower()
        if not (f_lower.endswith(".pdf") or f_lower.endswith(".doc") or f_lower.endswith(".docx")):
            continue

        full_path = os.path.join(folder_path, filename)
        # Determine if it's PDF or Word
        if f_lower.endswith(".pdf"):
            final_path = full_path
            # We'll parse as PDF
            parsed_text = extract_text_from_pdf(final_path)
        else:
            # Word doc => try conversion
            pdf_path = convert_word_to_pdf(full_path, converted_dir)
            if pdf_path is not None:
                # We got a PDF => parse that
                parsed_text = extract_text_from_pdf(pdf_path)
                final_path = pdf_path
            else:
                # Conversion failed => fallback docx direct parse
                parsed_text = extract_text_from_docx(full_path)
                final_path = full_path

        # Save the extracted text in parsed_docs
        base_name = os.path.splitext(os.path.basename(final_path))[0]
        output_file = os.path.join(parsed_dir, f"{base_name}_parsed.txt")
        with open(output_file, "w", encoding="utf-8") as f_out:
            f_out.write(parsed_text)
        
        print(f"Parsed text saved to: {output_file}")

if __name__ == "__main__":
    folder_path = "Lab05/"
    main(folder_path)
