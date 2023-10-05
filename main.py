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

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    # Creating the first line with invoice number.
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', align='L', ln=1)
    # second line with date is written.
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Date: {date}', align='L', ln=1)

    # Reading the excel into a dataframe.
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    header_columns = df.columns

    # Removing _ and Capitalizing the header
    header_columns = [item.replace('_', ' ').title() for item in header_columns]

    # Creating headers
    pdf.set_font(family='Times', style='B', size=10)
    pdf.cell(w=25, h=8, txt=header_columns[0], align='L', border=1)
    pdf.cell(w=60, h=8, txt=header_columns[1], align='L', border=1)
    pdf.cell(w=40, h=8, txt=header_columns[2], align='L', border=1)
    pdf.cell(w=30, h=8, txt=header_columns[3], align='L', border=1)
    pdf.cell(w=30, h=8, txt=header_columns[4], align='L', ln=1, border=1)

    # Looping through the df
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.cell(w=25, h=8, txt=str(row['product_id']), align='L', border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), align='L', border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), align='L', border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), align='L', border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), align='L', ln=1, border=1)

    # Writing out the pdf
    pdf.output(f'pdfs/{filename}.pdf')