import fitz  # PyMuPDF
import spacy
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import time


# Load spaCy model
nlp = spacy.load("en_core_web_sm")


def extract_text_from_pdf(pdf_path):
    # Extract text from a PDF file.
    text = ""
    doc = fitz.open(pdf_path)
    for page in doc:
        text += page.get_text()

    print(text)
    return text


def extract_entities(text):
    # Extract names, phone numbers, and emails from text using spaCy and regex.
    doc = nlp(text)
    names = []
    phone_numbers = []
    emails = []

    # Extract names using spaCy NER
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            names.append(ent.text)

    # Extract phone numbers using regex
    phone_pattern = re.compile(r'\b\d{10}\b')
    phone_numbers = phone_pattern.findall(text)

    # Extract emails using regex
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    emails = email_pattern.findall(text)

    return names, phone_numbers, emails


def save_to_csv(names, phone_numbers, emails, output_file):
    # Save extracted data to a CSV file.
    data = {
        'Name': names,
        'Phone Number': phone_numbers,
        'Email': emails
    }

    # Create DataFrame
    df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data.items()]))

    # Save to CSV
    script_folder = os.path.dirname(__file__)
    output_csv_folder = os.path.join(script_folder, 'Output_csv_files/')
    if not os.path.isdir(output_csv_folder):
        os.makedirs(output_csv_folder)
    df.to_csv(output_csv_folder + os.path.basename(output_file).split('/')[-1], index=False)


def extract_images(pdf_path):
    doc = fitz.open(pdf_path)
    for page_index in range(len(doc)):
        page = doc[page_index]
        image_list = page.get_images()
        if image_list:
            print(
                f"[+] Found a total of {len(image_list)} images in page {page_index}")
        else:
            print("[!] No images found on page", page_index)
            continue
        for image_index, img in enumerate(page.get_images(), start=1):
            # get the XREF of the image
            xref = img[0]

            # extract the image bytes
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]

            # get the image extension
            image_ext = base_image["ext"]

        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Identity, dpi=None,
                                  colorspace=fitz.csRGB, clip=None, alpha=True, annots=True)
            page = str(page)
            page = page.split('/')[-1]

            pix.save("sample_pdf_image"+str(page)+".png",)
    time.sleep(15)
    exit()


def main(pdf_path, output_csv):
    text = extract_text_from_pdf(pdf_path)
    names, phone_numbers, emails = extract_entities(text)
    save_to_csv(names, phone_numbers, emails, output_csv)
    extract_images(pdf_path)


def drop(event):
    entry.delete(0, tk.END)
    entry.insert(0, event.data)


def browse_files():
    filename = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)


def process_pdf():
    pdf_path = entry.get()
    if pdf_path:
        # Call your main function here
        main(pdf_path, pdf_path + " output.csv")
    else:
        print("Please select a PDF file")


root = TkinterDnD.Tk()
root.title("PDF Extractor")

label = tk.Label(root, text="Drag and drop a PDF file or click browse")
label.pack(pady=20)

entry = tk.Entry(root, width=50)
entry.pack()

button = tk.Button(root, text="Browse", command=browse_files)
button.pack(pady=10)

button = tk.Button(root, text="Process PDF", command=process_pdf)
button.pack(pady=10)

listbox = tk.Listbox(root)
listbox.pack(fill=tk.BOTH, expand=True)

entry.drop_target_register(DND_FILES)
entry.dnd_bind("<<Drop>>", drop)

root.mainloop()


if __name__ == "__main__":
    pdf_path = "preprints202304.0491.v1.pdf"
    output_csv = pdf_path + " output.csv"
    main(pdf_path, output_csv)
