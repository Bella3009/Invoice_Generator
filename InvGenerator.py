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
    # Split will give 2 variables therefore like this the variables will be set
    inv_nr, date = filename.split("-")
    
    # Input the data in the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{inv_nr}", ln=1)
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}")
    
    pdf.output(f"PDFs/{filename}.pdf")
