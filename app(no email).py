import os
import string
import pandas as pd
import jinja2
import pdfkit
from datetime import datetime
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams

if __name__ == "__main__":

    excel_file =  r'Y:\My Drive\Admin\Account\0 Receipt master record (NEW).xlsx'
    data1 = pd.read_excel(excel_file, sheet_name='Receipt Maker')
    i = 0
    data = pd.read_excel(excel_file, sheet_name='MASTER RECORD')
    pdf_dir = os.path.abspath("..\Debit Note backup Master File\DN SETTLED")
    for DN_Number in data1['DN Number']:
        if f"{DN_Number}.pdf" in os.listdir(pdf_dir):
            global path
            path = fr"Y:\My Drive\Admin\Debit Note backup Master File\DN SETTLED\{DN_Number}.pdf"
            text = extract_text(path, laparams=LAParams())
            print(f"{DN_Number}: DN found!")

            df = pd.read_excel(r'Y:\My Drive\Admin\Account Receipt Maker\class.xls')
            product = 1
            # Iterate over the values in the 'classes' column
            for value in df['classes']:
                if text.find(value) != -1:
                    product = value
                    break
            if product == 1:
                product == "UNKNOWN!!!!"
            while product[i] in string.ascii_letters or product[i] == " " or product[i] == "-" or product[i] == "\'" or product[i] == "&" or product[i] == ".":
                i += 1
                if len(product) == i:
                    break
            product = product[0:i]


            result = data[data['DN Number'] == DN_Number]
            bank_in_date = result['Bank in Date'].values[0]
            bank_in_date = pd.to_datetime(result['Bank in Date'])
            bank_in_date = bank_in_date.dt.strftime('%d/%m/%Y').values[0]
            ac_number = result['A/C No.'].values[0]
            insured = result['Client Name'].values[0]
            if len(insured) < 43:
                insured1 = insured
                insured2 = ""
                insured3 = ""
                insured4 = ""
                insured5 = ""
            elif len(insured) < 86 and len(insured) >= 43:
                insured1 = insured[:43]
                insured2 = insured[43:]
                insured3 = ""
                insured4 = ""
                insured5 = ""
            elif len(insured) < 129 and len(insured) >= 86:
                insured1 = insured[:43]
                insured2 = insured[43:86]
                insured3 = insured[86:]
                insured4 = ""
                insured5 = ""
            elif len(insured) < 172 and len(insured) >= 129:
                insured1 = insured[:43]
                insured2 = insured[43:86]
                insured3 = insured[86:129]
                insured4 = insured[129:]
                insured5 = ""
            else:
                insured1 = insured[:43]
                insured2 = insured[43:86]
                insured3 = insured[86:129]
                insured4 = insured[129:172]
                insured5 = insured[172:]
        # try:
            found = text.find("Mortgagee")
            found2 = text[found:].find("To")  
            d = text[found+18:found2+found]
            while d[0].isalpha() == False:
                d = d[1:]
            d = d.split("\n")
            i = 0
            try:
                PN = d[2]
                if PN[0].isdigit() == False and PN[0].isalpha() == False:
                    PN = "TBA"
            except:
                PN = "TBA"
            
            findcharge = text.find("Total Charges") 
            while text[findcharge].isdigit() == False:
                findcharge += 1
            findcharge2 = text[findcharge:].find("\n")
            dn_amount = text[findcharge:findcharge+findcharge2]

            context = { 'insured1': insured1, 'insured2': insured2, 'insured3': insured3, 'insured4': insured4, 'insured5': insured5, 'today_date': bank_in_date, 'total': f'${dn_amount}',
                        'subtotal1': dn_amount, 'product': product,
                        'DN_Number': DN_Number, 'AC_Number': ac_number, 'Policy_Number': PN,
                        }
            template_loader = jinja2.FileSystemLoader('./')
            template_env = jinja2.Environment(loader=template_loader)

            html_template = 'invoice.html'
            template = template_env.get_template(html_template)
            output_text = template.render(context)

            config = pdfkit.configuration(wkhtmltopdf=r'Y:\My Drive\Admin\Account Receipt Maker\wkhtmltopdf.exe')
            output_pdf = fr'Y:\My Drive\Admin\Account Receipt Maker\Receipts\{DN_Number}.pdf'
            pdfkit.from_string(output_text, output_pdf, configuration=config, css='invoice.css')
            # except Exception as error:
            #     print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            #     print(DN_Number + " failed!")
            #     print("An exception occurred:", type(error).__name__)
            #     print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        else:
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print(f"{DN_Number}: DN not found!")
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
