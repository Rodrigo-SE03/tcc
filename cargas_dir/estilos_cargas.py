import pandas as pd

merge_style ={
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
        "border_color": "white"
    }

merge_style2 = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
    }
    
empty_style = {
        "bold": 1,
        "right": 1,
        "fg_color": "#4f81bd",
        "border_color": "white"
    }

#Definição do estilo visual da tabela de consumo por carga
def consumo_equip_style(worksheet,workbook,grupo):
    global merge_style
    global empty_style
    merge_format = workbook.add_format(merge_style)
    empty_format = workbook.add_format(empty_style)
    if grupo == "Grupo A":
        worksheet.merge_range("C1:E1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("F1:H1","Consumo Diário Típico (kWh)",merge_format)
    else:
        worksheet.merge_range("C1:F1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("G1:J1","Consumo Diário Típico (kWh)",merge_format)
    worksheet.write_blank(0,0,'',empty_format)
    worksheet.write_blank(0,1,'',empty_format)
#--------------------------------------------------------------------------------------------------------

#Definição do estilo visual das tabelas de custos diários e mensais
def custos(custo,writer,workbook,worksheet,categoria,dias,tarifas_dict):
    global merge_style2
    merge_format = workbook.add_format(merge_style2)
    
    col = 7
    row = 2
    
    if categoria == 'Convencional':
        d_results = {'Consumo': [custo/tarifas_dict['convencional'],custo]}
        m_results = {'Consumo': [custo*dias/tarifas_dict['convencional'],custo*dias]}
    elif categoria == 'Branca':
        col += 1
        d_results = {
            'Consumo FP': [custo[0]/tarifas_dict['branca'][0],custo[0]],
            'Consumo I': [custo[1]/tarifas_dict['branca'][1],custo[1]],
            'Consumo P': [custo[2]/tarifas_dict['branca'][2],custo[2]]
        }
        m_results = {
            'Consumo FP': [custo[0]*dias/tarifas_dict['branca'][0],custo[0]*dias],
            'Consumo I': [custo[1]*dias/tarifas_dict['branca'][1],custo[1]*dias],
            'Consumo P': [custo[2]*dias/tarifas_dict['branca'][2],custo[2]*dias],
            'Total': [custo[0]*dias/tarifas_dict['branca'][0]+custo[1]*dias/tarifas_dict['branca'][1]+custo[2]*dias/tarifas_dict['branca'][2],(custo[0]+custo[1]+custo[2])*dias]
        }
    elif categoria == 'Verde':
        col += 3
        d_results = {
            'Consumo FP': [custo[0]/tarifas_dict['verde'][0],custo[0]],
            'Consumo P': [custo[1]/tarifas_dict['verde'][1],custo[1]],
            'Demanda': [custo[2]/tarifas_dict['verde'][2],custo[2]]
        }
        m_results = {
            'Consumo FP': [custo[0]*dias/tarifas_dict['verde'][0],custo[0]*dias],
            'Consumo P': [custo[1]*dias/tarifas_dict['verde'][1],custo[1]*dias],
            'Demanda': [custo[2]/tarifas_dict['verde'][2],custo[2]],
            'Total': ['-',(custo[0]+custo[1])*dias+custo[2]]
        }
    else:
        col += 3
        d_results = {
            'Consumo FP': [custo[0]/tarifas_dict['azul'][0],custo[0]],
            'Consumo P': [custo[1]/tarifas_dict['azul'][1],custo[1]],
            'Demanda FP': [custo[2]/tarifas_dict['azul'][2],custo[2]],
            'Demanda P': [custo[3]/tarifas_dict['azul'][3],custo[3]]
        }
        m_results = {
            'Consumo FP': [custo[0]*dias/tarifas_dict['azul'][0],custo[0]*dias],
            'Consumo P': [custo[1]*dias/tarifas_dict['azul'][1],custo[1]*dias],
            'Demanda FP': [custo[2]/tarifas_dict['azul'][2],custo[2]],
            'Demanda P': [custo[3]/tarifas_dict['azul'][3],custo[3]],
            'Total': ['-',(custo[0]+custo[1])*dias+custo[2]+custo[3]]
        }

    border = workbook.add_format({'border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    pot_format = workbook.add_format({'num_format':'#,##0.00 "kW"','border':1})
    nrg_format = workbook.add_format({'num_format':'#,##0.00 "kWh"','border':1})

    n_col = col
    n_row = row
    i=0
    for key in d_results.keys():
        worksheet.write(n_row,n_col,key,border)
        while i < len(m_results[key]):
            if i == 0 and n_row<row+2: format = nrg_format
            elif i == 1: format = rs_format
            elif i == 0 and n_row>=row+2 and categoria != "Branca" and d_results[key][i]!='-': format = pot_format
            elif i == 0 and n_row>=row+2 and categoria != "Branca" and d_results[key][i]=='-': format = border
            else: format=nrg_format
            worksheet.write(n_row,n_col+1,d_results[key][i],format)
            i+=1
            n_col+=1
        i=0
        n_col = col
        n_row+=1

    i=0
    n_col = col+5
    n_row = row
    print(n_col,n_row)
    for key in m_results.keys():
        worksheet.write(n_row,n_col,key,border)
        while i < len(m_results[key]):
            if i == 0 and n_row<row+2: format = nrg_format
            elif i == 1: format = rs_format
            elif i == 0 and n_row>=row+2 and categoria != "Branca" and m_results[key][i]!='-': format = pot_format
            elif i == 0 and n_row>=row+2 and categoria != "Branca" and m_results[key][i]=='-': format = border
            else: format=nrg_format
            worksheet.write(n_row,n_col+1,m_results[key][i],format)
            i+=1
            n_col+=1
        i=0
        n_col = col+5
        n_row+=1

    worksheet.merge_range(row-1,col,row-1,col+2,"Valores Diários",merge_format)
    worksheet.merge_range(row-1,col+5,row-1,col+5+2,"Valores Mensais",merge_format)
    return m_results
#--------------------------------------------------------------------------------------------------------

#Definição do estilo visual das tabelas de elementos reativos
def tabela_reativos(tabela_dict,writer,demr,categoria,consumo_mes):
    global merge_style2
    global merge_style
    global empty_style

    df = pd.DataFrame(tabela_dict)
    df.to_excel(writer,sheet_name=f'Reativos - {categoria}',startrow=3,header=False,index=False)
    workbook = writer.book
    worksheet = writer.sheets[f'Reativos - {categoria}']
    table_format = workbook.add_format(
    {
        'num_format': '#,##0.00'
    }
    )
    (max_row, max_col) = df.shape
    if categoria == "Verde":
        worksheet.set_column('A:J',10,table_format)
    else:
        worksheet.set_column('A:K',10,table_format)
    column_settings = [{"header": column} for column in df.columns]
    worksheet.add_table(2, 0, max_row+2, max_col - 1, {"columns": column_settings})
    merge_format = workbook.add_format(merge_style)
    merge_format2 = workbook.add_format(merge_style2)
    empty_format = workbook.add_format(empty_style)
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    worksheet.merge_range("A1:A2","",empty_format)

    if categoria == "Verde":
        worksheet.merge_range("B1:D1","Valores Ativos",merge_format)
        worksheet.merge_range("C2:D2","Consumo",merge_format)
        worksheet.merge_range("E1:F2","Valores Reativos",merge_format)
        worksheet.merge_range("G1:H2","Fator de Potência",merge_format)
        worksheet.merge_range("I1:J1","Valores Calculados",merge_format)
        worksheet.write("B2","Demanda",merge_format)
        worksheet.write("I2","Demanda",merge_format)
        worksheet.write("J2","Consumo",merge_format)
        worksheet.merge_range("M1:N1","Acréscimo na Fatura",merge_format2)
        worksheet.write("M2","Demanda",rs_format)
        worksheet.write("N2",demr,rs_format)
        worksheet.write("M3","Consumo",rs_format)
        worksheet.write("N3",consumo_mes,rs_format)
        worksheet.write("M4","Total",rs_format)
        worksheet.write("N4",consumo_mes+demr,rs_format)
    else:
        worksheet.merge_range("B1:E1","Valores Ativos",merge_format)
        worksheet.merge_range("B2:C2","Demanda",merge_format)
        worksheet.merge_range("D2:E2","Consumo",merge_format)
        worksheet.merge_range("F1:G2","Valores Reativos",merge_format)
        worksheet.merge_range("H1:I2","Fator de Potência",merge_format)
        worksheet.merge_range("J1:K1","Valores Calculados",merge_format)
        worksheet.write("J2","Demanda",merge_format)
        worksheet.write("K2","Consumo",merge_format)
        worksheet.merge_range("N1:O1","Acréscimo na Fatura",merge_format2)
        worksheet.write("N2","Demanda FP",rs_format)
        worksheet.write("O2",demr[0],rs_format)
        worksheet.write("N3","Demanda P",rs_format)
        worksheet.write("O3",demr[1],rs_format)
        worksheet.write("N4","Consumo",rs_format)
        worksheet.write("O4",consumo_mes,rs_format)
        worksheet.write("N5","Total",rs_format)
        worksheet.write("O5",consumo_mes+demr[0]+demr[1],rs_format)
         
    worksheet.autofit()
#--------------------------------------------------------------------------------------------------------

#Definição do estilo visual da tabela comparativa
def comparativo_style(grupo,comp_dict,writer):
    global merge_style2

    col = 3
    row = 4
    
    workbook = writer.book
    workbook.add_worksheet('Comparativo')
    worksheet = writer.sheets['Comparativo']
    merge_format2 = workbook.add_format(merge_style2)
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    header_format = workbook.add_format({
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    if grupo == 'Grupo B':
        worksheet.merge_range("C3:E3","Comparação de Custos",merge_format2)
    else:
        worksheet.merge_range("C3:I3","Comparação de Custos",merge_format2)

    i=0
    for key in comp_dict.keys():
        worksheet.write(3,i+2,key,header_format)
        worksheet.write_column(4,i+2,comp_dict[key],rs_format)
        i+=1
    worksheet.autofit()
    worksheet.set_column('F:F',14)
#--------------------------------------------------------------------------------------------------------