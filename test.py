import xlsxwriter
import pandas as pd
from PyPDF2 import PdfReader
import re


reader = PdfReader(f'fatura.pdf')
page = reader.pages[0]
text = page.extract_text()
text = re.findall(r'( ([DNOSAJMF][A-Z]+ [ \/0-9,A-Z]+\n){13})',text)[0]
print(text)
# text = text.replace('\n',' ')
# values = text.split(' ')
# print(values)