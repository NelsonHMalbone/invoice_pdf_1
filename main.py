import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem

    invoice_nr = filename.split('-')[0]
    invoice_date = filename.split('-')[1]

    # invoice number
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr: {invoice_nr}", align="L", ln=1)

    # creating date
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", align="L", ln=1)

    # creating table
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # adding a header to columns
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]

    pdf.set_font(family='Times', size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # adding the rows for table
    for index, row in df.iterrows():
        # used this to get the infomation from the table
        # print(row)
        pdf.set_font(family='Times', size=10)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # total price column
    total_sum = df["total_price"].sum()
    pdf.set_font(family='Times', size=10)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # total sum sentence
    pdf.set_font(family='Times', size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"The total amount due is: {total_sum}", ln=2)

    # company logo and name
    pdf.set_font(family='Times', size=12, style="B")
    pdf.cell(w=25, h=8, txt='PythonHow')
    pdf.image('pythonhow.png', w=6, h=6)

    # output of pdf
    pdf.output(f'PDF/{filename}.pdf')
