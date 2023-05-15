import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Retrieve all Excel files
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Retrieve the data for the PDF
    filename = Path(filepath).stem
    # Only the first part of the name is the invoice number
    inv_nr = filename.split("-")[0]
    # Input the data in the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{inv_nr}")
    pdf.output(f"PDFs/{filename}.pdf")

