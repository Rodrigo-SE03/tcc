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
    
    col = 8
    row = 2
    
    if categoria == 'Convencional':
        du_results = {'Consumo': [custo['Dias Úteis']/tarifas_dict['convencional'],custo['Dias Úteis']]}
        s_results = {'Consumo': [custo['Sábados']/tarifas_dict['convencional'],custo['Sábados']]}
        d_results = {'Consumo': [custo['Domingos']/tarifas_dict['convencional'],custo['Domingos']]}
        mdu_results = {'Consumo': [custo['Dias Úteis']*dias['dias_u']/tarifas_dict['convencional'],custo['Dias Úteis']*dias['dias_u']]}
        ms_results = {'Consumo': [custo['Sábados']*dias['dias_s']/tarifas_dict['convencional'],custo['Sábados']*dias['dias_s']]}
        md_results = {'Consumo': [custo['Domingos']*dias['dias_d']/tarifas_dict['convencional'],custo['Domingos']*dias['dias_d']]}

        dia_results = [du_results,s_results,d_results]
        mes_results = [mdu_results,ms_results,md_results]
        total_results = [{'Consumo': [mdu_results['Consumo'][0]+ms_results['Consumo'][0]+md_results['Consumo'][0],mdu_results['Consumo'][1]+ms_results['Consumo'][1]+md_results['Consumo'][1]]}]
    elif categoria == 'Branca':
        col += 1
        du_results = {
            'Consumo FP': [custo['Dias Úteis'][0]/tarifas_dict['branca'][0],custo['Dias Úteis'][0]],
            'Consumo I': [custo['Dias Úteis'][1]/tarifas_dict['branca'][1],custo['Dias Úteis'][1]],
            'Consumo P': [custo['Dias Úteis'][2]/tarifas_dict['branca'][2],custo['Dias Úteis'][2]],
            'Total': [custo['Dias Úteis'][2]/tarifas_dict['branca'][2]+custo['Dias Úteis'][1]/tarifas_dict['branca'][1]+custo['Dias Úteis'][0]/tarifas_dict['branca'][0]
                      ,custo['Dias Úteis'][0]+custo['Dias Úteis'][1]+custo['Dias Úteis'][2]]
        }
        s_results = {
            'Consumo FP': [custo['Sábados'][0]/tarifas_dict['branca'][0],custo['Sábados'][0]],
            'Consumo I': [custo['Sábados'][1]/tarifas_dict['branca'][1],custo['Sábados'][1]],
            'Consumo P': [custo['Sábados'][2]/tarifas_dict['branca'][2],custo['Sábados'][2]],
            'Total': [custo['Sábados'][2]/tarifas_dict['branca'][2]+custo['Sábados'][1]/tarifas_dict['branca'][1]+custo['Sábados'][0]/tarifas_dict['branca'][0]
                      ,custo['Sábados'][0]+custo['Sábados'][1]+custo['Sábados'][2]]
        }
        d_results = {
            'Consumo FP': [custo['Domingos'][0]/tarifas_dict['branca'][0],custo['Domingos'][0]],
            'Consumo I': [custo['Domingos'][1]/tarifas_dict['branca'][1],custo['Domingos'][1]],
            'Consumo P': [custo['Domingos'][2]/tarifas_dict['branca'][2],custo['Domingos'][2]],
            'Total': [custo['Domingos'][2]/tarifas_dict['branca'][2]+custo['Domingos'][1]/tarifas_dict['branca'][1]+custo['Domingos'][0]/tarifas_dict['branca'][0]
                      ,custo['Domingos'][0]+custo['Domingos'][1]+custo['Domingos'][2]]
        }
        mdu_results = {
            'Consumo FP': [custo['Dias Úteis'][0]*dias['dias_u']/tarifas_dict['branca'][0],custo['Dias Úteis'][0]*dias['dias_u']],
            'Consumo I': [custo['Dias Úteis'][1]*dias['dias_u']/tarifas_dict['branca'][1],custo['Dias Úteis'][1]*dias['dias_u']],
            'Consumo P': [custo['Dias Úteis'][2]*dias['dias_u']/tarifas_dict['branca'][2],custo['Dias Úteis'][2]*dias['dias_u']],
            'Total': [custo['Dias Úteis'][0]*dias['dias_u']/tarifas_dict['branca'][0]+custo['Dias Úteis'][1]*dias['dias_u']/tarifas_dict['branca'][1]+custo['Dias Úteis'][2]*dias['dias_u']/tarifas_dict['branca'][2],(custo['Dias Úteis'][0]+custo['Dias Úteis'][1]+custo['Dias Úteis'][2])*dias['dias_u']]
        }
        ms_results = {
            'Consumo FP': [custo['Sábados'][0]*dias['dias_s']/tarifas_dict['branca'][0],custo['Sábados'][0]*dias['dias_s']],
            'Consumo I': [custo['Sábados'][1]*dias['dias_s']/tarifas_dict['branca'][1],custo['Sábados'][1]*dias['dias_s']],
            'Consumo P': [custo['Sábados'][2]*dias['dias_s']/tarifas_dict['branca'][2],custo['Sábados'][2]*dias['dias_s']],
            'Total': [custo['Sábados'][0]*dias['dias_s']/tarifas_dict['branca'][0]+custo['Sábados'][1]*dias['dias_s']/tarifas_dict['branca'][1]+custo['Sábados'][2]*dias['dias_s']/tarifas_dict['branca'][2],(custo['Sábados'][0]+custo['Sábados'][1]+custo['Sábados'][2])*dias['dias_s']]
        }
        md_results = {
            'Consumo FP': [custo['Domingos'][0]*dias['dias_d']/tarifas_dict['branca'][0],custo['Domingos'][0]*dias['dias_d']],
            'Consumo I': [custo['Domingos'][1]*dias['dias_d']/tarifas_dict['branca'][1],custo['Domingos'][1]*dias['dias_d']],
            'Consumo P': [custo['Domingos'][2]*dias['dias_d']/tarifas_dict['branca'][2],custo['Domingos'][2]*dias['dias_d']],
            'Total': [custo['Domingos'][0]*dias['dias_d']/tarifas_dict['branca'][0]+custo['Domingos'][1]*dias['dias_d']/tarifas_dict['branca'][1]+custo['Domingos'][2]*dias['dias_d']/tarifas_dict['branca'][2],(custo['Domingos'][0]+custo['Domingos'][1]+custo['Domingos'][2])*dias['dias_d']]
        }

        dia_results = [du_results,s_results,d_results]
        mes_results = [mdu_results,ms_results,md_results]
        total_results = [{
            'Consumo FP': [mdu_results['Consumo FP'][0]+ms_results['Consumo FP'][0]+md_results['Consumo FP'][0],mdu_results['Consumo FP'][1]+ms_results['Consumo FP'][1]+md_results['Consumo FP'][1]],
            'Consumo I': [mdu_results['Consumo I'][0]+ms_results['Consumo I'][0]+md_results['Consumo I'][0],mdu_results['Consumo I'][1]+ms_results['Consumo I'][1]+md_results['Consumo I'][1]],
            'Consumo P': [mdu_results['Consumo P'][0]+ms_results['Consumo P'][0]+md_results['Consumo P'][0],mdu_results['Consumo P'][1]+ms_results['Consumo P'][1]+md_results['Consumo P'][1]],
            'Total': [mdu_results['Total'][0]+ms_results['Total'][0]+md_results['Total'][0],mdu_results['Total'][1]+ms_results['Total'][1]+md_results['Total'][1]]
        }]
    elif categoria == 'Verde':
        col += 3
        du_results = {
            'Consumo FP': [custo['Dias Úteis'][0]/tarifas_dict['verde'][0],custo['Dias Úteis'][0]],
            'Consumo P': [custo['Dias Úteis'][1]/tarifas_dict['verde'][1],custo['Dias Úteis'][1]],
            'Demanda': [custo['Dias Úteis'][2]/tarifas_dict['verde'][2],custo['Dias Úteis'][2]],
            'Total': [custo['Dias Úteis'][1]/tarifas_dict['verde'][1]+custo['Dias Úteis'][0]/tarifas_dict['verde'][0],
                      custo['Dias Úteis'][0]+custo['Dias Úteis'][1]]
        }
        s_results = {
            'Consumo FP': [custo['Sábados'][0]/tarifas_dict['verde'][0],custo['Sábados'][0]],
            'Consumo P': [custo['Sábados'][1]/tarifas_dict['verde'][1],custo['Sábados'][1]],
            'Demanda': ['-','-'],
            'Total': [custo['Sábados'][1]/tarifas_dict['verde'][1]+custo['Sábados'][0]/tarifas_dict['verde'][0],
                      custo['Sábados'][0]+custo['Sábados'][1]]
        }
        d_results = {
            'Consumo FP': [custo['Domingos'][0]/tarifas_dict['verde'][0],custo['Domingos'][0]],
            'Consumo P': [custo['Domingos'][1]/tarifas_dict['verde'][1],custo['Domingos'][1]],
            'Demanda': ['-','-'],
            'Total': [custo['Domingos'][1]/tarifas_dict['verde'][1]+custo['Domingos'][0]/tarifas_dict['verde'][0],
                      custo['Domingos'][0]+custo['Domingos'][1]]
        }
        
        mdu_results = {
            'Consumo FP': [custo['Dias Úteis'][0]*dias['dias_u']/tarifas_dict['verde'][0],custo['Dias Úteis'][0]*dias['dias_u']],
            'Consumo P': [custo['Dias Úteis'][1]*dias['dias_u']/tarifas_dict['verde'][1],custo['Dias Úteis'][1]*dias['dias_u']],
            'Demanda': [custo['Dias Úteis'][2]/tarifas_dict['verde'][2],custo['Dias Úteis'][2]],
            'Total': ['-',(custo['Dias Úteis'][0]+custo['Dias Úteis'][1])*dias['dias_u']+custo['Dias Úteis'][2]]
        }
        ms_results = {
            'Consumo FP': [custo['Sábados'][0]*dias['dias_s']/tarifas_dict['verde'][0],custo['Sábados'][0]*dias['dias_s']],
            'Consumo P': [custo['Sábados'][1]*dias['dias_s']/tarifas_dict['verde'][1],custo['Sábados'][1]*dias['dias_s']],
            'Demanda': ['-','-'],
            'Total': ['-',(custo['Sábados'][0]+custo['Sábados'][1])*dias['dias_s']]
        }
        md_results = {
            'Consumo FP': [custo['Domingos'][0]*dias['dias_d']/tarifas_dict['verde'][0],custo['Domingos'][0]*dias['dias_d']],
            'Consumo P': [custo['Domingos'][1]*dias['dias_d']/tarifas_dict['verde'][1],custo['Domingos'][1]*dias['dias_d']],
            'Demanda': ['-','-'],
            'Total': ['-',(custo['Domingos'][0]+custo['Domingos'][1])*dias['dias_d']]
        }

        mes_results = [mdu_results,ms_results,md_results]
        dia_results = [du_results,s_results,d_results]
        total_results = [{
            'Consumo FP': [mdu_results['Consumo FP'][0]+ms_results['Consumo FP'][0]+md_results['Consumo FP'][0],mdu_results['Consumo FP'][1]+ms_results['Consumo FP'][1]+md_results['Consumo FP'][1]],
            'Consumo P': [mdu_results['Consumo P'][0]+ms_results['Consumo P'][0]+md_results['Consumo P'][0],mdu_results['Consumo P'][1]+ms_results['Consumo P'][1]+md_results['Consumo P'][1]],
            'Demanda': [custo['Dias Úteis'][2]/tarifas_dict['verde'][2],custo['Dias Úteis'][2]],
            'Total': [mdu_results['Total'][0]+ms_results['Total'][0]+md_results['Total'][0],mdu_results['Total'][1]+ms_results['Total'][1]+md_results['Total'][1]]
        }]
    else:
        col += 3
        du_results = {
            'Consumo FP': [custo['Dias Úteis'][0]/tarifas_dict['azul'][0],custo['Dias Úteis'][0]],
            'Consumo P': [custo['Dias Úteis'][1]/tarifas_dict['azul'][1],custo['Dias Úteis'][1]],
            'Demanda FP': [custo['Dias Úteis'][2]/tarifas_dict['azul'][2],custo['Dias Úteis'][2]],
            'Demanda P': [custo['Dias Úteis'][3]/tarifas_dict['azul'][2],custo['Dias Úteis'][3]],
            'Total': [custo['Dias Úteis'][1]/tarifas_dict['azul'][1]+custo['Dias Úteis'][0]/tarifas_dict['azul'][0],
                      custo['Dias Úteis'][0]+custo['Dias Úteis'][1]]
        }
        s_results = {
            'Consumo FP': [custo['Sábados'][0]/tarifas_dict['azul'][0],custo['Sábados'][0]],
            'Consumo P': [custo['Sábados'][1]/tarifas_dict['azul'][1],custo['Sábados'][1]],
            'Demanda P': ['-','-'],
            'Demanda FP': ['-','-'],
            'Total': [custo['Sábados'][1]/tarifas_dict['azul'][1]+custo['Sábados'][0]/tarifas_dict['azul'][0],
                      custo['Sábados'][0]+custo['Sábados'][1]]
        }
        d_results = {
            'Consumo FP': [custo['Domingos'][0]/tarifas_dict['azul'][0],custo['Domingos'][0]],
            'Consumo P': [custo['Domingos'][1]/tarifas_dict['azul'][1],custo['Domingos'][1]],
            'Demanda P': ['-','-'],
            'Demanda FP': ['-','-'],
            'Total': [custo['Domingos'][1]/tarifas_dict['azul'][1]+custo['Domingos'][0]/tarifas_dict['azul'][0],
                      custo['Domingos'][0]+custo['Domingos'][1]]
        }

        mdu_results = {
            'Consumo FP': [custo['Dias Úteis'][0]*dias['dias_u']/tarifas_dict['azul'][0],custo['Dias Úteis'][0]*dias['dias_u']],
            'Consumo P': [custo['Dias Úteis'][1]*dias['dias_u']/tarifas_dict['azul'][1],custo['Dias Úteis'][1]*dias['dias_u']],
            'Demanda FP': [custo['Dias Úteis'][2]/tarifas_dict['azul'][2],custo['Dias Úteis'][2]],
            'Demanda P': [custo['Dias Úteis'][3]/tarifas_dict['azul'][3],custo['Dias Úteis'][3]],
            'Total': ['-',(custo['Dias Úteis'][0]+custo['Dias Úteis'][1])*dias['dias_u']+custo['Dias Úteis'][2]+custo['Dias Úteis'][3]]
        }
        ms_results = {
            'Consumo FP': [custo['Sábados'][0]*dias['dias_s']/tarifas_dict['azul'][0],custo['Sábados'][0]*dias['dias_s']],
            'Consumo P': [custo['Sábados'][1]*dias['dias_s']/tarifas_dict['azul'][1],custo['Sábados'][1]*dias['dias_s']],
            'Demanda P': ['-','-'],
            'Demanda FP': ['-','-'],
            'Total': ['-',(custo['Sábados'][0]+custo['Sábados'][1])*dias['dias_s']]
        }
        md_results = {
            'Consumo FP': [custo['Domingos'][0]*dias['dias_d']/tarifas_dict['azul'][0],custo['Domingos'][0]*dias['dias_d']],
            'Consumo P': [custo['Domingos'][1]*dias['dias_d']/tarifas_dict['azul'][1],custo['Domingos'][1]*dias['dias_d']],
            'Demanda FP': ['-','-'],
            'Demanda P': ['-','-'],
            'Total': ['-',(custo['Domingos'][0]+custo['Domingos'][1])*dias['dias_d']]
        }

        mes_results = [mdu_results,ms_results,md_results]
        dia_results = [du_results,s_results,d_results]
        total_results = [{
            'Consumo FP': [mdu_results['Consumo FP'][0]+ms_results['Consumo FP'][0]+md_results['Consumo FP'][0],mdu_results['Consumo FP'][1]+ms_results['Consumo FP'][1]+md_results['Consumo FP'][1]],
            'Consumo P': [mdu_results['Consumo P'][0]+ms_results['Consumo P'][0]+md_results['Consumo P'][0],mdu_results['Consumo P'][1]+ms_results['Consumo P'][1]+md_results['Consumo P'][1]],
            'Demanda FP': [custo['Dias Úteis'][2]/tarifas_dict['azul'][2],custo['Dias Úteis'][2]],
            'Demanda P': [custo['Dias Úteis'][3]/tarifas_dict['azul'][3],custo['Dias Úteis'][3]],
            'Total': [mdu_results['Total'][0]+ms_results['Total'][0]+md_results['Total'][0],mdu_results['Total'][1]+ms_results['Total'][1]+md_results['Total'][1]]
        }]

    border = workbook.add_format({'border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    pot_format = workbook.add_format({'num_format':'#,##0.00 "kW"','border':1})
    nrg_format = workbook.add_format({'num_format':'#,##0.00 "kWh"','border':1})
    dias_format = workbook.add_format({'border':1,'bold':1,'fg_color':'#D9D9D9',"align": "center","valign": "vcenter"})

    n_col = col
    n_row = row
    i=0
    j=0
    for result in dia_results:
        for key in result.keys():
            worksheet.write(n_row,n_col,key,border)
            while i < len(result[key]):
                if i == 0 and n_row<row+2: format = nrg_format
                elif i == 1: format = rs_format
                elif i == 0 and n_row>=row+2 and categoria != "Branca" and result[key][i]!='-': format = pot_format
                elif i == 0 and n_row>=row+2 and categoria != "Branca" and result[key][i]=='-': format = border
                else: format=nrg_format
                worksheet.write(n_row,n_col+1,result[key][i],format)
                i+=1
                n_col+=1
            i=0
            n_col = col+3*j
            n_row+=1
        j+=1
        n_col = col+3*j
        n_row = row
        

    i=0
    j=0
    row_m = 4 + len(mes_results[0])
    n_col = col
    n_row = row_m
    for result in mes_results:
        for key in result.keys():
            worksheet.write(n_row,n_col,key,border)
            while i < len(result[key]):
                if i == 0 and n_row<row_m+2: format = nrg_format
                elif i == 1: format = rs_format
                elif i == 0 and n_row>=row_m+2 and categoria != "Branca" and result[key][i]!='-': format = pot_format
                elif i == 0 and n_row>=row_m+2 and categoria != "Branca" and result[key][i]=='-': format = border
                else: format=nrg_format
                worksheet.write(n_row,n_col+1,result[key][i],format)
                i+=1
                n_col+=1
            i=0
            n_col = col+3*j
            n_row+=1
        j+=1
        n_col = col+3*j
        n_row = row_m
    

    i=0
    j=0
    row_t = row-1
    n_col = col+12
    n_row = row_t
    for result in total_results:
        for key in result.keys():
            worksheet.write(n_row,n_col,key,border)
            while i < len(result[key]):
                if i == 0 and n_row<row_t+2: format = nrg_format
                elif i == 1: format = rs_format
                elif i == 0 and n_row>=row_t+2 and categoria != "Branca" and result[key][i]!='-': format = pot_format
                elif i == 0 and n_row>=row_t+2 and categoria != "Branca" and result[key][i]=='-': format = border
                else: format=nrg_format
                worksheet.write(n_row,n_col+1,result[key][i],format)
                i+=1
                n_col+=1
            i=0
            n_col = col+12+3*j
            n_row+=1
        j+=1
        n_col = col+3*j
        n_row = row_t

    worksheet.merge_range(row-2,col,row-2,col+8,"Valores Diários",merge_format)
    worksheet.merge_range(row-1,col,row-1,col+2,"Dias Úteis",dias_format)
    worksheet.merge_range(row-1,col+3,row-1,col+5,"Sábados",dias_format)
    worksheet.merge_range(row-1,col+6,row-1,col+8,"Domingos",dias_format)

    worksheet.merge_range(row_m-2,col,row_m-2,col+8,"Valores Mensais",merge_format)
    worksheet.merge_range(row_m-1,col,row_m-1,col+2,"Dias Úteis",dias_format)
    worksheet.merge_range(row_m-1,col+3,row_m-1,col+5,"Sábados",dias_format)
    worksheet.merge_range(row_m-1,col+6,row_m-1,col+8,"Domingos",dias_format)

    worksheet.merge_range(row_t-1,col+12,row_t-1,col+14,"Total Mensal",merge_format)

    m_results=total_results[0]
    return m_results
#--------------------------------------------------------------------------------------------------------

#Criação da curva de consumo diária
def criar_grafico(worksheet,workbook,categoria,dias): 
    chart_du = workbook.add_chart({'type':'column'})
    if dias['dias_s'] != 0:
        chart_s = workbook.add_chart({'type':'column'})
    if dias['dias_d'] != 0:
        chart_d = workbook.add_chart({'type':'column'})

    if categoria == 'Convencional':
        chart_du.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
        if dias['dias_s'] != 0:
            chart_s.add_series({'categories':f"='Consumo - {categoria}'!$C$1444:$C$2883",'name': "Potência",'values':f"='Consumo - {categoria}'!$D$1444:$D$2883"})
        if dias['dias_d'] != 0:
            chart_d.add_series({'categories':f"='Consumo - {categoria}'!$C$2886:$C$4324",'name': "Potência",'values':f"='Consumo - {categoria}'!$D$2886:$D$4324"})
    elif categoria == 'Branca':
        chart_du.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
        chart_du.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2:$E$1441"})
        chart_du.add_series({'name':"Potência - Intermediário",'values':f"='Consumo - {categoria}'!$F$2:$F$1441"})
        if dias['dias_s'] != 0:
            chart_s.add_series({'categories':f"='Consumo - {categoria}'!$C$1444:$C$2883",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$1444:$D$2883"})
            chart_s.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$1444:$E$2883"})
            chart_s.add_series({'name':"Potência - Intermediário",'values':f"='Consumo - {categoria}'!$F$1444:$F$2883"})
        if dias['dias_d'] != 0:
            chart_d.add_series({'categories':f"='Consumo - {categoria}'!$C$2886:$C$4324",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2886:$D$4324"})
            chart_d.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2886:$E$4324"})
            chart_d.add_series({'name':"Potência - Intermediário",'values':f"='Consumo - {categoria}'!$F$2886:$F$4324"})
    else:
        chart_du.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
        chart_du.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2:$E$1441"})
        line_chart_du = workbook.add_chart({'type':'line'})
        line_chart_du.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "FP",'values':f"='Consumo - {categoria}'!$H$2:$H$1441","y2_axis":True,'line':{'color':'red','width':1.5}})
        line_chart_du.add_series({'name':"Limite - FP",'values':f"='Consumo - {categoria}'!$I$2:$I$1441","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
        line_chart_du.add_series({'name':"Lim2",'values':f"='Consumo - {categoria}'!$J$2:$J$1441","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
        line_chart_du.set_y2_axis({'name':'Fator de Potência'})
        chart_du.combine(line_chart_du)
        if dias['dias_s'] != 0:
            chart_s.add_series({'categories':f"='Consumo - {categoria}'!$C$1444:$C$2883",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$1444:$D$2883"})
            chart_s.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$1444:$E$2883"})
            line_chart_s = workbook.add_chart({'type':'line'})
            line_chart_s.add_series({'categories':f"='Consumo - {categoria}'!$C$1444:$C$2883",'name': "FP",'values':f"='Consumo - {categoria}'!$H$1444:$H$2883","y2_axis":True,'line':{'color':'red','width':1.5}})
            line_chart_s.add_series({'name':"Limite - FP",'values':f"='Consumo - {categoria}'!$I$1444:$I$2883","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
            line_chart_s.add_series({'name':"Lim2",'values':f"='Consumo - {categoria}'!$J$1444:$J$2883","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
            line_chart_s.set_y2_axis({'name':'Fator de Potência'})
            chart_s.combine(line_chart_s)
        if dias['dias_d'] != 0:
            chart_d.add_series({'categories':f"='Consumo - {categoria}'!$C$2886:$C$4324",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2886:$D$4324"})
            chart_d.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2886:$E$4324"})
            line_chart_d = workbook.add_chart({'type':'line'})
            line_chart_d.add_series({'categories':f"='Consumo - {categoria}'!$C$2886:$C$4324",'name': "FP",'values':f"='Consumo - {categoria}'!$H$2886:$H$4324","y2_axis":True,'line':{'color':'red','width':1.5}})
            line_chart_d.add_series({'name':"Limite - FP",'values':f"='Consumo - {categoria}'!$I$2886:$I$4324","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
            line_chart_d.add_series({'name':"Lim2",'values':f"='Consumo - {categoria}'!$J$2886:$J$4324","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
            line_chart_d.set_y2_axis({'name':'Fator de Potência'})
            chart_d.combine(line_chart_d)
    
    chart_du.set_x_axis(
    {
        "interval_unit": 60,
        "num_format": "h",
        'name':'Tempo - horas'
    })
    chart_du.set_legend({'position':'bottom','delete_series':[4]})
    chart_du.set_y_axis({'name':'Potência - kW'})
    chart_du.set_size({'width': 860, 'height': 450})
    chart_du.set_title({'name':'Perfil de Consumo - Dias Úteis'})

    if dias['dias_s'] != 0:
        chart_s.set_x_axis(
        {
            "interval_unit": 60,
            "num_format": "h",
            'name':'Tempo - horas'
        })
        chart_s.set_legend({'position':'bottom','delete_series':[4]})
        chart_s.set_y_axis({'name':'Potência - kW'})
        chart_s.set_size({'width': 860, 'height': 450})
        chart_s.set_title({'name':'Perfil de Consumo - Sábados'})

    if dias['dias_d'] != 0:
        chart_d.set_x_axis(
        {
            "interval_unit": 60,
            "num_format": "h",
            'name':'Tempo - horas'
        })
        chart_d.set_legend({'position':'bottom','delete_series':[4]})
        chart_d.set_y_axis({'name':'Potência - kW'})
        chart_d.set_size({'width': 860, 'height': 450})
        chart_d.set_title({'name':'Perfil de Consumo - Domingos'})

    if categoria == "Branca":
        worksheet.insert_chart('H14', chart_du)
        if dias['dias_s'] != 0:
            worksheet.insert_chart('S14', chart_s)
        if dias['dias_d'] != 0:
            worksheet.insert_chart('AE14', chart_d)
    elif categoria == 'Verde' or categoria == 'Azul':
        worksheet.insert_chart('L15', chart_du)
        if dias['dias_s'] != 0:
            worksheet.insert_chart('T15', chart_s)
        if dias['dias_d'] != 0:
            worksheet.insert_chart('AF15', chart_d)
    else:
        chart_du.set_legend({'none': True})
        worksheet.insert_chart('F9', chart_du)
        if dias['dias_s'] != 0:
            worksheet.insert_chart('Q9', chart_s)
        if dias['dias_d'] != 0:
            worksheet.insert_chart('AB9', chart_d)
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
        worksheet.write("J2","Energia",merge_format)
        worksheet.merge_range("M1:N1","Acréscimo na Fatura",merge_format2)
        worksheet.write("M2","DMCR",rs_format)
        worksheet.write("N2",demr,rs_format)
        worksheet.write("M3","UFER",rs_format)
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
        worksheet.write("K2","Energia",merge_format)
        worksheet.merge_range("N1:O1","Acréscimo na Fatura",merge_format2)
        worksheet.write("N2","DMCR FP",rs_format)
        worksheet.write("O2",demr[0],rs_format)
        worksheet.write("N3","DMCR P",rs_format)
        worksheet.write("O3",demr[1],rs_format)
        worksheet.write("N4","UFER",rs_format)
        worksheet.write("O4",consumo_mes,rs_format)
        worksheet.write("N5","Total",rs_format)
        worksheet.write("O5",consumo_mes+demr[0]+demr[1],rs_format)
         
    worksheet.autofit()
#--------------------------------------------------------------------------------------------------------

#Definição do estilo visual da tabela comparativa
def comparativo_style(grupo,comp_dict,pct_dict,writer):
    global merge_style2
    
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
    pct_format = workbook.add_format({'num_format':'0.00%','border':1,"align": "center"})

    chart = workbook.add_chart({'type':'column'})
    if grupo == 'Grupo B':
        worksheet.merge_range("C3:D3","Comparação de Custos",merge_format2)
        chart.add_series({'categories':f"=Comparativo!$D$4",'name': "Convencional",'values':f"=Comparativo!$D$5",'fill':{'color':'#F2EA50'},'overlap':-20})
        chart.add_series({'name': "Branca",'values':f"=Comparativo!$D$6",'fill':{'color':'#D9D9D9'}})
        chart.add_series({'name': "Diferença",'values':f"=Comparativo!$D$7",'fill':{'color':'#ED5D5D'}})
        chart.set_size({'width': 410, 'height': 440})
    else:
        
        worksheet.merge_range("C3:I3","Comparação de Custos",merge_format2)
        worksheet.merge_range("C9:H9","Representação Percentual",merge_format2)
        chart.add_series({'categories':f"=Comparativo!$E$4:$G$4",'name': "Verde",'values':f"=Comparativo!$E$5:$G$5",'fill':{'color':'#00B050'},'overlap':-20})
        chart.add_series({'name': "Azul",'values':f"=Comparativo!$E$6:$G$6",'fill':{'color':'#0070C0'}})
        chart.add_series({'name': "Diferença",'values':f"=Comparativo!$E$7:$G$7",'fill':{'color':'#ED5D5D'}})
        chart.set_size({'width': 860, 'height': 450})

    chart.set_y_axis({'num_format': "R$ #,##0.00"})
    chart.set_title({'name':'Comparação de Custos'})
    chart.set_legend({'position': 'right'})

    i=0
    for key in comp_dict.keys():
        worksheet.write(3,i+2,key,header_format)
        worksheet.write_column(4,i+2,comp_dict[key],rs_format)
        i+=1
    
    i=0
    if grupo == 'Grupo A':
        worksheet.insert_chart('C14', chart)
        for key in pct_dict.keys():
            worksheet.write(9,i+2,key,header_format)
            worksheet.write_column(10,i+2,pct_dict[key],pct_format)
            i+=1
    else:
        worksheet.insert_chart('C11', chart)
        worksheet.write(7,3,comp_dict['Total'][3],pct_format)
    worksheet.autofit()
    worksheet.set_column('F:F',14)
    worksheet.set_column('L:Q',14)
#--------------------------------------------------------------------------------------------------------