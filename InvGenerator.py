import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Retrieve all Excel files
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
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
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Set the table header
    excel_columns = df.columns
    excel_columns = [item.replace("_", " ").title() for item in excel_columns]
    
    pdf.set_font(family="Times", size=16)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8,txt=excel_columns[0], border=1)
    pdf.cell(w=70, h=8,txt=excel_columns[1],border=1)
    pdf.cell(w=30, h=8,txt=excel_columns[2], border=1)
    pdf.cell(w=30, h=8,txt=excel_columns[3], border=1)
    pdf.cell(w=30, h=8,txt=excel_columns[4], border=1, ln=1)
    
    # Set row data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=16)
        pdf.set_text_color(80, 80, 80)
        # Use the string function because integer value cause an error
        pdf.cell(w=30, h=8,txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8,txt=str(row["product_name"]),border=1)
        pdf.cell(w=30, h=8,txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8,txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8,txt=str(row["total_price"]), border=1, ln=1)
        
    # Total row
    total = df["total_price"].sum()
    pdf.set_font(family="Times", size=16)
    pdf.cell(w=30, h=8,txt=f"The total price is {total}", ln=1)
    
    # Add logo and company name
    pdf.set_font(family="Times", size=16)
    pdf.cell(w=30, h=8,txt="PythonHow")
    pdf.image("pythonhow.png", w=10)
    
    pdf.output(f"PDFs/{filename}.pdf")
