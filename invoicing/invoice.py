import os
import pandas as pd
import glob
# glob is used to search files and paths
from fpdf import FPDF
from pathlib import Path
# pathlib is also used to read paths of files

def generate(invoices_path,pdfs_path,product_id,product_name,amount_purchased,
             price_per_unit,total_price,company_name=None,image_path=None):
    """
        This function converts invoices Excel files into PDF invoices.
        :param invoices_path: Path where the invoice Excel files are located.
        :param pdfs_path: Path where the generated PDF files will be saved.
        :param product_id: Column name for product ID in the Excel files.
        :param product_name: Column name for product name in the Excel files.
        :param amount_purchased: Column name for amount purchased in the Excel files.
        :param price_per_unit: Column name for price per unit in the Excel files.
        :param total_price: Column name for total price in the Excel files.
        :param image_path: (Optional) Path to the image to include in the PDF.
        :return: None
        """
    filepaths = glob.glob(f"{invoices_path}/*.xlsx")
    # it will search all the files with texts after *

    for filepath in filepaths:
        df = pd.read_excel(filepath,sheet_name="Sheet 1")

        pdf = FPDF(orientation="P",unit="mm",format="A4")
        pdf.add_page()

        filename = Path(filepath).stem
        # Path(filepath).stem  it will read the stems of the files
        invoice_nr = filename.split("-")[0]
        date = filename.split("-")[1]

        pdf.set_font(family="Times",size=16,style="B")
        pdf.cell(w=50,h=8, txt=f"Invoice number.{invoice_nr}",ln=1)

        pdf.set_font(family="Times",size=16,style="B")
        pdf.cell(w=50,h=8, txt=f"Date-{date}",ln=1)

        columns = df.columns
        # df.columns shows all the columns of a file,whereas df.rows will show the rows under these columns
        columns = [item.replace("_"," ").title()for item in columns]
        pdf.set_font(family="Times",size=10)
        pdf.cell(w=30, h=8, txt=columns[0], border=1)
        pdf.cell(w=70, h=8, txt=columns[1], border=1)
        pdf.cell(w=30, h=8, txt=columns[2], border=1)
        pdf.cell(w=30, h=8, txt=columns[3], border=1)
        pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

        for index,row in df.iterrows():
            pdf.cell(w=30, h=8, txt=str(row[product_id]),border=1)
            pdf.cell(w=70, h=8, txt=str(row[product_name]),border=1)
            pdf.cell(w=30, h=8, txt=str(row[amount_purchased]),border=1)
            pdf.cell(w=30, h=8, txt=str(row[price_per_unit]),border=1)
            pdf.cell(w=30, h=8, txt=str(row[total_price]),border=1,ln=1)

        total_sum = df["total_price"].sum()
        pdf.set_font(family="Times",size=10)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=70, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30,h=8,txt=f"The total price is {total_sum}",ln=1)

        pdf.set_font(family="Times", size=10,style="B")
        if company_name:
            pdf.cell(w=25, h=8, txt=company_name)
        if image_path:
            pdf.image(image_path,  x=5, y=285, w=10)


        os.makedirs(pdfs_path, exist_ok=True)#if already exist for that exist_ok is used
        pdf.output(f"{pdfs_path}/{filename}.pdf")
        # if we put this outside the loop only the last file will get the changes not all the files