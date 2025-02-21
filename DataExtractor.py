import re
import csv
import fitz  # pymupdf
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_text_from_pdf(pdf_filename):
    """extracts text from a pdf, grouping pages in pairs as property data"""
    try:
        doc = fitz.open(pdf_filename)
        total_pages = len(doc)
        extracted_text = ""
        property_count = 1

        for page_num in range(2, total_pages, 2):
            extracted_text += f"=== property {property_count} ===\n"
            text_1 = doc[page_num].get_text("text")
            text_2 = doc[page_num + 1].get_text("text") if page_num + 1 < total_pages else ""
            extracted_text += text_1 + "\n" + text_2 + "\n" + "=" * 80 + "\n"
            property_count += 1

        return extracted_text
    except Exception as e:
        messagebox.showerror("error", f"something went wrong: {e}")
        return ""

def extract_value(text, keyword, pattern=r"(\S.+)"):
    """finds the keyword and extracts the value using regex"""
    match = re.search(rf"{keyword}\s*\n{pattern}", text, re.IGNORECASE)
    return match.group(1).strip() if match else "N/A"

def extract_helaa_ref(text):
    """finds and returns the first occurrence of hxxxx format"""
    match = re.search(r"\bH\d{4}\b", text)
    return match.group(0).strip() if match else "N/A"

def extract_conclusion_comments(text):
    """finds the 'conclusion' section and extracts the following 'comments'"""
    match = re.search(r"Conclusion.*?Comments\s*\n([\s\S]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else "N/A"

def process_extracted_text(text, output_csv):
    """processes extracted text and writes formatted data to csv"""
    headers = [
        "location", "address", "owner / estate agent", "contact details",
        "initial date of contact", "latest date of contact", "status",
        "planning permission", "helaa reference", "conclusion comments",
        "site size (ha)", "est dwellings"
    ]

    properties = re.split(r"=== property \d+ ===", text)[1:]
    data = []

    for prop in properties:
        location = extract_value(prop, "Parish")
        address = extract_value(prop, "Site Address")
        helaa_ref = extract_helaa_ref(prop)
        site_size = extract_value(prop, "Site Size \(Hectares\)")
        conclusion_comments = extract_conclusion_comments(prop)
        conclusion_comments = re.sub(r"\n={10,}", "", conclusion_comments).strip()
        est_dwellings = extract_value(prop, "Included Capacity \(dwellings\)")
        est_dwellings = est_dwellings if est_dwellings.isdigit() else "N/A"

        data.append([
            location, address, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", helaa_ref, conclusion_comments, site_size, est_dwellings
        ])

    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(data)

    messagebox.showinfo("success", f"data extraction complete! saved as '{output_csv}'.")

def main():
    """gui for file selection"""
    root = tk.Tk()
    root.withdraw()  # hide main window

    messagebox.showinfo("pdf to csv converter", "select the pdf file to process.")

    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        messagebox.showwarning("no file selected", "you must select a pdf file.")
        return

    output_csv = pdf_file.rsplit(".", 1)[0] + "_formatted.csv"
    extracted_text = extract_text_from_pdf(pdf_file)

    if extracted_text:
        process_extracted_text(extracted_text, output_csv)

if __name__ == "__main__":
    main()
