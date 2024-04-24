import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split('-')[0]
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr: {invoice_nr}", align="L", ln=1, border=1)
    pdf.output(f'PDF/{filename}.pdf')
