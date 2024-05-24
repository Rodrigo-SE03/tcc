import xlsxwriter
import pandas as pd
from datetime import datetime


dem_p = []
dem_fp = []
dem_r = []
con_p = []
con_fp = []
con_r = []
hr_con = []
hr_r = []

df = pd.read_excel('Fatura - modelo.xlsx')
try:
    if df.iloc[0,9] == 'Validar':
        pass
    else:
        print('fail')
except:
    print('fail')

m0 = df['Mês'].tolist()[0]
data = [0,0]
data[0] = m0.split('/')[0]
data[1] = f'20{m0.split('/')[1]}'
print(data)