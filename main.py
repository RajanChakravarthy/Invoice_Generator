import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# list of all the xlsx files is created
filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    # Extracting invoice number and date from filepath
    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    df = pd.read_excel(filepath)

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', align='L', ln=1)

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Date: {date}', align='L', ln=1)


    # Writing out the pdf
    pdf.output(f'pdfs/{filename}.pdf')