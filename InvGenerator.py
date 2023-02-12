import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    # other option which is the same is 
    # date = filename.split("-")
  
    pdf.set_font(family="Times", size=16, style="B", ln=1)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=2)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    colum = df.columns
    colum = [item.replace("-", " ").title() for item in colum]
    for n in range(4):
        pdf.set_font(family="Times", size=12, style="B")
        pdf.cell(w=40, h=8, txt=colum[n], border=1)
     
     pdf.cell(w=40, h=8, txt=colum[4], border=1, ln=1)
       
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.cell(w=40, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["total_price"]), border=1, ln=1)
    
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=12)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt=str(total_sum), border=1, ln=1)
    
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=0, h=8, txt=f"The total price is {total_sum}", ln=1)
    
    pdf.cell(w=40, h=8, txt="Company name")
    pdf.image("pythonhow.png", w=10)
    
    pdf.output(f"PDFs/{filename}.pdf")
