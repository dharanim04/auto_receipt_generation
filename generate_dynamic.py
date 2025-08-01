import pandas as pd
from docx import Document
import os
from tkinter import Tk, filedialog
from docx2pdf import convert
from datetime import datetime, date
from data_folder.send_email import send_email_with_attachment
from data_folder.helper_file import get_body

body_file_path = get_body()

# --- Select Excel File ---
print('Upload excel file-- for Data')
Tk().withdraw()
excel_file = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not excel_file:
    print("❌ No Excel file selected.")
    exit()

# --- Select Word Template ---
print('Upload word file for template')
template_file = filedialog.askopenfilename(
    title="Select Word Template (.docx)",
    filetypes=[("Word Documents", "*.docx *.doc")]
)
if not template_file:
    print("❌ No Word template selected.")
    exit()

# -------Input PDF or Word ---------
# isPdf = input('Do you want OutFile PDF format? Y/N')

# --- Load Data ---
df = pd.read_excel(excel_file)
headers = df.columns.tolist()

# --- Output Directory ---
output_dir_word = "generated_docs"
os.makedirs(output_dir_word, exist_ok=True)

output_dir_pdf = "generated_pdf"
os.makedirs(output_dir_pdf, exist_ok=True)

def format_value(value):
    try:
        if value == 'nan' or pd.isna(value):
            return ""
        elif isinstance(value, date):
            val= datetime.date(value)
            return str(val)
        else:
            return str(value)
    except Exception as e:
            print(f"Failed to convert value", value)
            return ""
   

def replace_placeholders(doc, row):
    for paragraph in doc.paragraphs:
        i = 0
        while i < len(paragraph.runs):
            text_accum = ""
            run_indices = []
            j = i
            while j < len(paragraph.runs) and len(run_indices) < 10:  # Prevent runaway loop
                text_accum += paragraph.runs[j].text
                run_indices.append(j)
                for key in headers:
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in text_accum:
                        replaced = text_accum.replace(placeholder, format_value(row[key]))
                        # Clear and replace in first run
                        paragraph.runs[run_indices[0]].text = replaced
                        # Clear the remaining runs
                        for k in run_indices[1:]:
                            paragraph.runs[k].text = ""
                        i = run_indices[-1] + 1  # Move index past replaced runs
                        break
                else:
                    j += 1
                    continue
                break
            else:
                i += 1

def generate_docs(row):
    doc = Document(template_file)
    replace_placeholders(doc, row)

    name = row.get("Name", f"Record_{row.name}")
    name_date = f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    file_name_with_date = "".join(c if c.isalnum() else "_" for c in str(name_date))

    docx_path = os.path.join(output_dir_word, f"{file_name_with_date}.docx")
    doc.save(docx_path)

    # if isPdf.lower() == 'y':
    pdf_path = os.path.join(output_dir_pdf, f"{file_name_with_date}.pdf")
    try:
        convert(docx_path, pdf_path)
    except AssertionError as e:
        print("❌ AssertionError:", e)
        print("💡 Is Microsoft Word installed and activated?")
    except Exception as e:
        print(f"⚠️ PDF conversion failed for {file_name_with_date}: {e}")
    
    # load body
    body=''
    try:
        with open(body_file_path, 'r', encoding='utf-8') as f:
            body = f.read()
    except FileNotFoundError:
        print(f"Error: File not found at '{body_file_path}'")

      # Send email with attachment (customize recipient and message as needed)
    to_email = row.get('Email', None)  # Assumes 'Email' column exists
    if to_email and not pd.isna(to_email) and str(to_email).strip() :
        subject = f"Donation Receipt - {file_name_with_date}"
        body_msg = f"Dear {name}, \n \n{body}"
        try:
            send_email_with_attachment(to_email, subject, body_msg, pdf_path)
        except Exception as e:
            print(f"Failed to send email to {to_email}: {e}")
    else:
        print('Sender email address missing')
# --- Generate for All Rows ---
for _, row in df.iterrows():
    generate_docs(row)

print("✅ All Word and PDF files generated in:", output_dir_word, output_dir_pdf)
