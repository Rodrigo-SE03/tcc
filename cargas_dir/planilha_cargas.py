import pandas as pd
import copy
import os
import math
from cargas_dir import estilos_cargas

#Função geral de criação de planilhas
def criar_planilha(cargas,tarifas_dict,grupo,nome,folder,h_p,dias):
    cargas_dict = copy.deepcopy(cargas)
    cargas_dict.pop('Remover')
    writer = pd.ExcelWriter(f'{folder}/{nome}',engine="xlsxwriter")
    tab_cargas(cargas_dict=cargas_dict,writer=writer)
    equip_dict = tab_consumo_por_carga(cargas=cargas_dict,writer=writer,grupo=grupo,h_p=h_p)
    if grupo == 'Grupo B':
        valores_C = tab_consumo(categoria='Convencional',itens=cargas_dict,writer=writer,h_p=h_p,equip_dict=equip_dict,tarifas_dict=tarifas_dict,dias=dias)
        valores_B = tab_consumo(categoria='Branca',itens=cargas_dict,writer=writer,h_p=h_p,equip_dict=equip_dict,tarifas_dict=tarifas_dict,dias=dias)
        comparativo_gpb(m_results_C=valores_C[1],m_results_B=valores_B[1],grupo=grupo,writer=writer)
    else:    
        valores_V = tab_consumo(categoria='Verde',itens=cargas_dict,writer=writer,h_p=h_p,equip_dict=equip_dict,tarifas_dict=tarifas_dict,dias=dias)
        valores_A = tab_consumo(categoria='Azul',itens=cargas_dict,writer=writer,h_p=h_p,equip_dict=equip_dict,tarifas_dict=tarifas_dict,dias=dias)
        reativos_V = tab_reativos(categoria='Verde',consumo_dict=valores_V[0],h_p=h_p,tarifas_dict=tarifas_dict,writer=writer,dias=dias)
        reativos_A = tab_reativos(categoria='Azul',consumo_dict=valores_A[0],h_p=h_p,tarifas_dict=tarifas_dict,writer=writer,dias=dias)
        comparativo_gpa(m_results_A=valores_A[1],m_results_V=valores_V[1],r_results_A=reativos_A,r_results_V=reativos_V,grupo=grupo,writer=writer)
    
    writer.close()
#--------------------------------------------------------------------------------------------------------

#Criação da aba de cargas
def tab_cargas(cargas_dict,writer):
    df_cargas = pd.DataFrame(cargas_dict)
    df_cargas.to_excel(writer, sheet_name="Cargas", startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Cargas"]
    (max_row, max_col) = df_cargas.shape
    column_settings = [{"header": column} for column in df_cargas.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    worksheet.autofit()
#--------------------------------------------------------------------------------------------------------

#Função para calcular horas no horário de ponta e fora de ponta
def calc_intervalo(inicio,fim,h_p,grupo,postos=False):
    ponta = [*range(int(h_p*60),int((h_p+3)*60))]
    inter = [*range(int((h_p-1)*60),int(h_p*60)),*range(int((h_p+3)*60),int((h_p+4)*60))]

    if grupo == 'Grupo A':
        fora = [*range(0,int(h_p*60)),*range(int((h_p+3)*60),24*60)]
    else:
        fora = [*range(0,int((h_p-1)*60)),*range(int((h_p+4)*60),24*60)]

    hr_i = int(inicio.split(":")[0])*60 + int(inicio.split(":")[1])
    hr_f = int(fim.split(":")[0])*60 + int(fim.split(":")[1])
    hrs = [*range(hr_i,hr_f)]

    p=0
    i=0
    fp=0

    for h in hrs:
        if h in ponta:
            p += 1
        if h in inter:
            i += 1
        if h in fora:
            fp += 1
    
    p = p/60
    i = i/60
    fp = fp/60

    if postos:
        return[fora,inter,ponta]
    elif grupo == 'Grupo A':
        return [fp,p]
    else:
        return [fp,i,p]
#--------------------------------------------------------------------------------------------------------

#Criação da aba de consumo por carga
def tab_consumo_por_carga(cargas,writer,grupo,h_p):
    if grupo == 'Grupo A':
        equip_dict = {"Carga":[],
                    "Potência (kW)":[],
                    "H - Ponta":[],
                    "H - Fora Ponta":[],
                    "Total - H":[],
                    "C - Ponta":[],
                    "C - Fora Ponta":[],
                    "Total - C":[]}
    else:
        equip_dict = {"Carga":[],
                    "Potência (kW)":[],
                    "H - Ponta":[],
                    "H - Intermediário":[],
                    "H - Fora Ponta":[],
                    "Total - H":[],
                    "C - Ponta":[],
                    "C - Intermediário":[],
                    "C - Fora Ponta":[],
                    "Total - C":[]}
    
    i=0
    for carga in cargas['Carga']:
        equip_dict['Carga'].append(carga)
        equip_dict['Potência (kW)'].append(cargas['Potência (kW)'][i])
        if grupo == 'Grupo A':
            equip_dict['H - Ponta'].append(calc_intervalo(inicio=cargas['Início'][i],fim=cargas['Fim'][i],h_p=h_p,grupo=grupo)[1])
            equip_dict['H - Fora Ponta'].append(calc_intervalo(inicio=cargas['Início'][i],fim=cargas['Fim'][i],h_p=h_p,grupo=grupo)[0])
            equip_dict['Total - H'].append(equip_dict['H - Ponta'][i]+equip_dict['H - Fora Ponta'][i])
            equip_dict['C - Ponta'].append(equip_dict['Potência (kW)'][i]*equip_dict['H - Ponta'][i])
            equip_dict['C - Fora Ponta'].append(equip_dict['Potência (kW)'][i]*equip_dict['H - Fora Ponta'][i])
            equip_dict['Total - C'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total - H'][i])
        else:
            equip_dict['H - Ponta'].append(calc_intervalo(inicio=cargas['Início'][i],fim=cargas['Fim'][i],h_p=h_p,grupo=grupo)[2])
            equip_dict['H - Intermediário'].append(calc_intervalo(inicio=cargas['Início'][i],fim=cargas['Fim'][i],h_p=h_p,grupo=grupo)[1])
            equip_dict['H - Fora Ponta'].append(calc_intervalo(inicio=cargas['Início'][i],fim=cargas['Fim'][i],h_p=h_p,grupo=grupo)[0])
            equip_dict['Total - H'].append(equip_dict['H - Ponta'][i]+equip_dict['H - Fora Ponta'][i]+equip_dict['H - Intermediário'][i])
            equip_dict['C - Ponta'].append(equip_dict['Potência (kW)'][i]*equip_dict['H - Ponta'][i])
            equip_dict['C - Fora Ponta'].append(equip_dict['Potência (kW)'][i]*equip_dict['H - Fora Ponta'][i])
            equip_dict['C - Intermediário'].append(equip_dict['Potência (kW)'][i]*equip_dict['H - Intermediário'][i])
            equip_dict['Total - C'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total - H'][i])
        i+=1

    df_equip = pd.DataFrame(equip_dict)
    df_equip.to_excel(writer, sheet_name="Consumo por carga", startrow=2, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo por carga"]
    (max_row, max_col) = df_equip.shape
    column_settings = [{"header": column} for column in df_equip.columns]
    worksheet.add_table(1, 0, max_row+1, max_col-1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    estilos_cargas.consumo_equip_style(worksheet,workbook,grupo)
    worksheet.autofit()
    return equip_dict
#--------------------------------------------------------------------------------------------------------

#Função que retorna o valor de horas em decimal a partir de uma entrada no formato "hh:mm"
def get_hora(tempo):
    h = float(tempo.split(":")[0])
    m = float(tempo.split(":")[1])
    hora = h + (m/60)
    return hora
#--------------------------------------------------------------------------------------------------------

#Função para criar a tabela de consumo de cada modalidade tarifária
def select_consumo(itens,categoria,h_p):  
    h_ponta = h_p
    grupo = 'Grupo B' if categoria == 'Convencional' or categoria == 'Branca' else 'Grupo A'

    postos = calc_intervalo(inicio="00:00",fim="01:00",grupo=grupo,h_p=h_p,postos=True)

    if categoria == 'Convencional':
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência - kW':[]}
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot=0
                while i < len(itens['Carga']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):
                        pot += itens['Potência (kW)'][i]
                    i+=1
                consumo_dict['Potência - kW'].append(pot)
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m)
                consumo_dict['Instante'].append(0) 
    
    elif categoria == 'Branca':
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência FP - kW':[],'Potência P - kW':[],'Potência I - kW':[]}
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot_fp = 0
                pot_p = 0
                pot_i = 0
                while i < len(itens['Carga']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):  
                        if (h*60+m) in postos[1]:
                            pot_i += itens['Potência (kW)'][i]
                        elif (h*60+m) in postos[0]:
                            pot_fp += itens['Potência (kW)'][i]
                        else:
                            pot_p += itens['Potência (kW)'][i]
                    i+=1
                consumo_dict['Potência FP - kW'].append(pot_fp)
                consumo_dict['Potência P - kW'].append(pot_p)
                consumo_dict['Potência I - kW'].append(pot_i)
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m)
                consumo_dict['Instante'].append(0)
    
    else:
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência FP - kW':[],'Potência P - kW':[],'Potência Reativa FP - kVAr':[],'Potência Reativa P - kVAr':[],'FP':[],'Limite - Indutivo':[],'Limite - Capacitivo':[]}
        j=0
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot_fp = 0
                pot_p = 0
                potr_fp = 0
                potr_p = 0
                while i < len(itens['Carga']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):
                        if (h*60+m) in postos[0]:
                            pot_fp += itens['Potência (kW)'][i]*itens['Quantidade'][i]
                            potr_fp += itens['Potência (kW)'][i]*math.sqrt((1/math.pow(itens['FP'][i],2))-1)*itens['Quantidade'][i] * (1 if itens['FP - Tipo'][i] == "Indutivo" else -1)
                        else:
                            pot_p += itens['Potência (kW)'][i]*itens['Quantidade'][i]
                            potr_p += itens['Potência (kW)'][i]*math.sqrt((1/math.pow((itens['FP'][i]),2))-1)*itens['Quantidade'][i] * (1 if itens['FP - Tipo'][i] == "Indutivo" else -1)
                    i+=1
                consumo_dict['Potência FP - kW'].append(pot_fp)
                consumo_dict['Potência P - kW'].append(pot_p)
                consumo_dict['Potência Reativa FP - kVAr'].append(potr_fp)
                consumo_dict['Potência Reativa P - kVAr'].append(potr_p)
                if h<h_ponta or h>=(h_ponta+3):
                    consumo_dict['FP'].append((consumo_dict['Potência FP - kW'][j]/math.sqrt(math.pow(consumo_dict['Potência FP - kW'][j],2)+math.pow(consumo_dict['Potência Reativa FP - kVAr'][j],2))) * (-1 if consumo_dict['Potência Reativa FP - kVAr'][j] < 0 else 1))
                else:
                    consumo_dict['FP'].append((consumo_dict['Potência P - kW'][j]/math.sqrt(math.pow(consumo_dict['Potência P - kW'][j],2)+math.pow(consumo_dict['Potência Reativa P - kVAr'][j],2))) * (-1 if consumo_dict['Potência Reativa P - kVAr'][j] < 0 else 1))
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m) 
                consumo_dict['Instante'].append(0) 
                consumo_dict['Limite - Capacitivo'].append(-0.92)
                consumo_dict['Limite - Indutivo'].append(0.92)
                j+=1
    
    return consumo_dict
#--------------------------------------------------------------------------------------------------------

#Função para calcular o custo da energia
def calc_custo(tarifas_dict,equip_dict,categoria,consumo_dict):
    if categoria == 'Convencional': custo = 0
    elif categoria == 'Branca': custo = [0,0,0]
    elif categoria == 'Verde':
        custo = [0,0,0]

        #Cálculo da demanda (valor máximo da média das potências medidas em 15 minutos)
        demanda = []
        i = 0
        d = 0 
        while i<len(consumo_dict['Potência FP - kW']):
            if (i % 15) == 0:
                demanda.append(d/15)
                d = 0
            d += consumo_dict['Potência FP - kW'][i] + consumo_dict['Potência P - kW'][i]
            i+=1
        demanda = max(demanda)
        #--------------------------------------------------------------------------------------------------------

        custo[2] = demanda*tarifas_dict['verde'][2]
    else: 
        custo = [0,0,0,0]

        #Cálculo da demanda fp (valor máximo da média das potências medidas em 15 minutos)
        demanda_fp = []
        i = 0
        d = 0 
        while i<len(consumo_dict['Potência FP - kW']):
            if (i % 15) == 0:
                demanda_fp.append(d/15)
                d = 0
            d += consumo_dict['Potência FP - kW'][i]
            i+=1
        demanda_fp = max(demanda_fp)
        #--------------------------------------------------------------------------------------------------------
        
        #Cálculo da demanda p (valor máximo da média das potências medidas em 15 minutos)
        demanda_p = []
        i = 0
        d = 0 
        while i<len(consumo_dict['Potência P - kW']):
            if (i % 15) == 0:
                demanda_p.append(d/15)
                d = 0
            d += consumo_dict['Potência P - kW'][i]
            i+=1
        demanda_p = max(demanda_p)
        #--------------------------------------------------------------------------------------------------------
        
        custo[2] = demanda_fp*tarifas_dict['azul'][2]
        custo[3] = demanda_p*tarifas_dict['azul'][3]
    i=0
    while i < len(equip_dict['Carga']):
        if categoria=='Convencional':
            custo += equip_dict['Total - C'][i]*tarifas_dict['convencional']
        elif categoria=='Branca':
            custo[0] += equip_dict['C - Fora Ponta'][i]*tarifas_dict['branca'][0]
            custo[1] += equip_dict['C - Intermediário'][i]*tarifas_dict['branca'][1]
            custo[2] += equip_dict['C - Ponta'][i]*tarifas_dict['branca'][2]
        elif categoria == 'Verde':
            custo[0] += equip_dict['C - Fora Ponta'][i]*tarifas_dict['verde'][0]
            custo[1] += equip_dict['C - Ponta'][i]*tarifas_dict['verde'][1]
        else:
            custo[0] += equip_dict['C - Fora Ponta'][i]*tarifas_dict['azul'][0]
            custo[1] += equip_dict['C - Ponta'][i]*tarifas_dict['azul'][1]
        i+=1

    return custo

#--------------------------------------------------------------------------------------------------------

#Criação da aba de consumo geral
def tab_consumo(itens,writer,categoria,h_p,equip_dict,tarifas_dict,dias):  
    consumo_dict = select_consumo(itens,categoria,h_p)

    custo = calc_custo(consumo_dict=consumo_dict,categoria=categoria,equip_dict=equip_dict,tarifas_dict=tarifas_dict)

    df_consumo = pd.DataFrame(consumo_dict)
    df_consumo.to_excel(writer, sheet_name=f"Consumo - {categoria}", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets[f"Consumo - {categoria}"]
    (max_row, max_col) = df_consumo.shape
    column_settings = [{"header": column} for column in df_consumo.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    custos_mensais = estilos_cargas.custos(custo=custo,writer=writer,categoria=categoria,workbook=workbook,worksheet=worksheet,dias=dias,tarifas_dict=tarifas_dict)

    hora_format = workbook.add_format({'num_format': 'hh:mm:ss'})
    i=1
    while i <= len(consumo_dict['Horas']):
        worksheet.write_formula(f'C{i+1}', f'=DATE(YEAR(TODAY()), MONTH(TODAY()), DAY(TODAY())) + TIME(A{i+1}, B{i+1}, 0)',hora_format)
        i+=1
    
    worksheet.set_column("A:B",None, None,{"hidden":True})
    worksheet.autofit()
    if categoria != 'Convencional':
        worksheet.set_column('J:K',16)
    if categoria == 'Verde' or 'Azul':
        worksheet.set_column('Q:Q',16)
    criar_grafico(worksheet,workbook,categoria)
    return [consumo_dict,custos_mensais]
#--------------------------------------------------------------------------------------------------------

#Criação da curva de consumo diária
def criar_grafico(worksheet,workbook,categoria): 
    chart = workbook.add_chart({'type':'column'})
    if categoria == 'Convencional':
        chart.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
    elif categoria == 'Branca':
        chart.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
        chart.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2:$E$1441"})
        chart.add_series({'name':"Potência - Intermediário",'values':f"='Consumo - {categoria}'!$F$2:$F$1441"})
    else:
        chart.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':f"='Consumo - {categoria}'!$D$2:$D$1441"})
        chart.add_series({'name':"Potência - Ponta",'values':f"='Consumo - {categoria}'!$E$2:$E$1441"})
        line_chart = workbook.add_chart({'type':'line'})
        line_chart.add_series({'categories':f"='Consumo - {categoria}'!$C$2:$C$1441",'name': "FP",'values':f"='Consumo - {categoria}'!$H$2:$H$1441","y2_axis":True,'line':{'color':'red','width':1.5}})
        line_chart.add_series({'name':"Limite - FP",'values':f"='Consumo - {categoria}'!$I$2:$I$1441","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
        line_chart.add_series({'name':"Lim2",'values':f"='Consumo - {categoria}'!$J$2:$J$1441","y2_axis":True,'line':{'color':'#92D050','width':1,'dash_type': 'long_dash'}})
        line_chart.set_y2_axis({'name':'Fator de Potência'})
        chart.combine(line_chart)
    
    chart.set_x_axis(
    {
        "interval_unit": 60,
        "num_format": "h",
        'name':'Tempo - horas'
    })
    chart.set_legend({'position':'bottom','delete_series':[4]})
    chart.set_y_axis({'name':'Potência - kW'})
    chart.set_size({'width': 860, 'height': 450})
    chart.set_title({'name':'Perfil de Consumo'})
    if categoria == "Branca":
        worksheet.insert_chart('J9', chart)
    elif categoria == 'Verde' or categoria == 'Azul':
        worksheet.insert_chart('L9', chart)
    else:
        chart.set_legend({'none': True})
        worksheet.insert_chart('I9', chart)
#--------------------------------------------------------------------------------------------------------

#Criação da aba com os elementos reativos identificados na instalação   
def tab_reativos(categoria,consumo_dict,h_p,tarifas_dict,writer,dias):
    demr=0
    demanda = max(consumo_dict['Potência FP - kW']) if max(consumo_dict['Potência FP - kW']) > max(consumo_dict['Potência P - kW']) else max(consumo_dict['Potência P - kW'])
    demanda_fp = max(consumo_dict['Potência FP - kW'])
    demanda_p = max(consumo_dict['Potência P - kW'])
    te_r = tarifas_dict['te']
    td_r = tarifas_dict['verde'][2]

    c_fp = [*range(0,24)]
    c_p = [*range(0,24)]
    cr_fp = [*range(0,24)]
    cr_p = [*range(0,24)]
    
    dem = [*range(0,24)]
    dem_fp = [*range(0,24)]
    dem_p = [*range(0,24)]
    dem_kw = [*range(0,24)]

    ind = [*range(0,24)]
    cap = [*range(0,24)]
    rs = [*range(0,24)]

    demr_p = 0
    demr_fp = 0

    i=0
    while i < len(rs):
        c_fp[i] = 0
        c_p[i] = 0
        cr_fp[i] = 0
        cr_p[i] = 0
        
        dem[i] = 0
        dem_fp[i] = 0
        dem_p[i] = 0
        dem_kw[i] = 0

        ind[i] = 0
        cap[i] = 0
        rs[i] = 0
        i+=1

    periodo=[]
    for h in range(0,24):
        periodo.append(f'{h}-{h+1}')

    j=0
    i=0
    consumo = 0
    reativo = 0
    maior = 0
    maior_fp = 0
    maior_p = 0
    for p in periodo:
        intervalo = [*range(int(p.split('-')[0])*60,int(p.split('-')[1])*60)]
        i = intervalo[0]
        while i <= intervalo[-1]:
            c_fp[j] += consumo_dict['Potência FP - kW'][i]
            c_p[j] += consumo_dict['Potência P - kW'][i]
            cr_fp[j] += consumo_dict['Potência Reativa FP - kVAr'][i]
            cr_p[j] += consumo_dict['Potência Reativa P - kVAr'][i]
            if categoria == 'Verde':
                dem[j] = consumo_dict['Potência FP - kW'][i] if consumo_dict['Potência FP - kW'][i] > dem[j] else dem[j]
                dem[j] = consumo_dict['Potência P - kW'][i] if consumo_dict['Potência P - kW'][i] > dem[j] else dem[j]
            else:
                dem_fp[j] = consumo_dict['Potência FP - kW'][i] if consumo_dict['Potência FP - kW'][i] > dem_fp[j] else dem_fp[j]
                dem_p[j] = consumo_dict['Potência P - kW'][i] if consumo_dict['Potência P - kW'][i] > dem_p[j] else dem_p[j]
            i+=1

        c_fp[j] = c_fp[j]/60
        c_p[j] = c_p[j]/60
        cr_fp[j] = cr_fp[j]/60
        cr_p[j] = cr_p[j]/60

        consumo = c_fp[j] if c_fp[j] != 0 else c_p[j]
        reativo = cr_fp[j] if cr_fp[j] != 0 else cr_p[j]
        if consumo == 0:
            ind[j] = 0
            cap[j] = 0
            dem_kw[j] = 0
            rs[j] = 0
        elif reativo < 0:
            cap[j] = consumo/math.sqrt(pow(consumo,2)+pow(reativo,2))
            ind[j] = 0
            if categoria == 'Verde':
                dem_kw[j] = dem[j]*(0.92/cap[j])
            else:
                dem_kw[j] = (0.92/cap[j])*(dem_fp[j] if int(p.split('-')[0]) < h_p or int(p.split('-')[0]) >= (h_p+3) else dem_p[j])
        elif reativo>0:
            cap[j] = 0
            ind[j] = consumo/math.sqrt(pow(consumo,2)+pow(reativo,2))
            if categoria == 'Verde':
                dem_kw[j] = dem[j]*(0.92/ind[j])
            else:
                dem_kw[j] = (0.92/ind[j])*(dem_fp[j] if int(p.split('-')[0]) < h_p or int(p.split('-')[0]) >= (h_p+3) else dem_p[j])
        else:
            cap[j] = 1
            ind[j] = 1
            if categoria == 'Verde':
                dem_kw[j] = dem[j]*(0.92/ind[j])
            else:
                dem_kw[j] = (0.92/ind[j])*(dem_fp[j] if int(p.split('-')[0]) < h_p or int(p.split('-')[0]) >= (h_p+3) else dem_p[j])
        
        if ind[j] !=0 or cap[j] != 0:
            if ind[j] > 0.92 or cap[j] >0.92:
                rs[j] = 0
            elif cap[j] !=0 and cap[j] < 0.92 and j>=6 and j<=17: rs[j] = 0
            else:
                rs[j] = math.fabs(consumo) * ((0.92/(ind[j] if ind[j] != 0 else cap[j]))-1) * te_r
        
        if dem_kw[j] > maior and categoria == 'Verde':
                maior = dem_kw[j]
                if ind[j] != 0 or cap[j] != 0:
                    if ind[j] >= 0.92 or cap[j] >= 0.92: demr = 0
                    elif cap[j] !=0 and cap[j] < 0.92 and j>=6 and j<=17: demr = 0
                    else: demr = (maior-demanda)*td_r
        
        if dem_kw[j] > maior_fp and dem_fp[j] != 0 and categoria == 'Azul':
                maior_fp = dem_kw[j]
                if ind[j] != 0 or cap[j] != 0:
                    if ind[j] >= 0.92 or cap[j] >= 0.92: demr_fp = 0
                    elif cap[j] !=0 and cap[j] < 0.92 and j>=6 and j<=17: demr_fp = 0
                    else: demr_fp = (maior_fp-demanda_fp)*td_r

        if dem_kw[j] > maior_p and dem_p[j] != 0 and categoria == 'Azul':
            maior_p = dem_kw[j]
            if ind[j] != 0 or cap[j] != 0:
                if ind[j] > 0.92 or cap[j] > 0.92: demr_p = 0
                elif cap[j] !=0 and cap[j] < 0.92 and j>=6 and j<=17: demr_p = 0
                else: demr_p = (maior_p-demanda_p)*td_r

        j+=1
    
    

    tabela_dict = { "Período":periodo,
                    "Fora Ponta - kW": dem_fp,
                    "Ponta - kW": dem_p,
                    "Fora Ponta - kWh": c_fp,
                    "Ponta - kWh": c_p,
                    "Fora Ponta - kVArh": cr_fp,
                    "Ponta - kVArh": cr_p,
                    "Indutivo": ind,
                    "Capacitivo": cap,
                    " kW": dem_kw,
                    "R$": rs
                    }
    if categoria == 'Verde':
        tabela_dict = { "Período":periodo,
                    "kW": dem,
                    "Fora Ponta - kWh": c_fp,
                    "Ponta - kWh": c_p,
                    "Fora Ponta - kVArh": cr_fp,
                    "Ponta - kVArh": cr_p,
                    "Indutivo": ind,
                    "Capacitivo": cap,
                    " kW": dem_kw,
                    "R$": rs
                    }
    if categoria == 'Azul':
        demr = [demr_fp,demr_p]
    consumo_mes = dias*sum(tabela_dict['R$'])
    estilos_cargas.tabela_reativos(categoria=categoria,demr=demr,tabela_dict=tabela_dict,writer=writer,consumo_mes=consumo_mes)
    return [consumo_mes,demr]
#--------------------------------------------------------------------------------------------------------

#Criação da aba de comparativo dos valores calculados para modalidades do grupo B
def comparativo_gpb(m_results_C,m_results_B,grupo,writer):
    custo_final_C = m_results_C['Consumo'][1]

    custo_final_B = m_results_B['Total'][1]

    comp_dict = {
        'Modalidade': ['Convencional','Branca','Diferença','Diferença Percentual'],
        # 'Consumo': [custo_final_C,custo_final_B,abs(custo_final_B-custo_final_C)],
        'Total': [custo_final_C,custo_final_B,abs(custo_final_B-custo_final_C),1-min([custo_final_B,custo_final_C])/max([custo_final_B,custo_final_C])]
    }
    estilos_cargas.comparativo_style(grupo=grupo,comp_dict=comp_dict,writer=writer,pct_dict={})
#--------------------------------------------------------------------------------------------------------

#Criação da aba de comparativo dos valores calculados para modalidades do grupo A
def comparativo_gpa(m_results_V,m_results_A,r_results_V,r_results_A,grupo,writer):
    custo_consumo_fp_V = m_results_V['Consumo FP'][1]
    custo_consumo_p_V = m_results_V['Consumo P'][1]
    custo_demanda_V = m_results_V['Demanda'][1]

    custo_consumo_fp_A = m_results_A['Consumo FP'][1]
    custo_consumo_p_A = m_results_A['Consumo P'][1]
    custo_demanda_fp_A = m_results_A['Demanda FP'][1]
    custo_demanda_p_A = m_results_A['Demanda P'][1]

    custo_consumo_r_V = r_results_V[0]
    custo_demanda_r_V = r_results_V[1]

    custo_consumo_r_A = r_results_A[0]
    custo_demanda_rfp_A = r_results_A[1][0]
    custo_demanda_rp_A = r_results_A[1][1]

    comp_dict = {
        'Modalidade': ['Verde','Azul','Diferença'],
        'Consumo FP': [custo_consumo_fp_V,custo_consumo_fp_A,abs(custo_consumo_fp_V-custo_consumo_fp_A)],
        'Consumo P': [custo_consumo_p_V,custo_consumo_p_A,abs(custo_consumo_p_V-custo_consumo_p_A)],
        'Demanda': [custo_demanda_V,custo_demanda_fp_A+custo_demanda_p_A,abs(custo_demanda_V-(custo_demanda_fp_A+custo_demanda_p_A))],
        'DMCR': [custo_demanda_r_V,custo_demanda_rfp_A+custo_demanda_rp_A,abs(custo_demanda_r_V-(custo_demanda_rfp_A+custo_demanda_rp_A))],
        'UFER': [custo_consumo_r_V,custo_consumo_r_A,abs(custo_consumo_r_V-custo_consumo_r_A)],
        'Total':[custo_consumo_fp_V+custo_consumo_p_V+custo_consumo_r_V+custo_demanda_r_V+custo_demanda_V,custo_consumo_fp_A+custo_consumo_p_A+custo_demanda_fp_A+custo_demanda_p_A+custo_consumo_r_A+custo_demanda_rfp_A+custo_demanda_rp_A,abs(custo_consumo_fp_V+custo_consumo_p_V+custo_consumo_r_V+custo_demanda_r_V+custo_demanda_V-(custo_consumo_fp_A+custo_consumo_p_A+custo_demanda_fp_A+custo_demanda_p_A+custo_consumo_r_A+custo_demanda_rfp_A+custo_demanda_rp_A))]
    }

    pct_dict = {
        'Modalidade': ['Verde','Azul'],
        'Consumo FP': [comp_dict['Consumo FP'][0]/comp_dict['Total'][0],comp_dict['Consumo FP'][1]/comp_dict['Total'][1]],
        'Consumo P': [comp_dict['Consumo P'][0]/comp_dict['Total'][0],comp_dict['Consumo P'][1]/comp_dict['Total'][1]],
        'Demanda': [comp_dict['Demanda'][0]/comp_dict['Total'][0],comp_dict['Demanda'][1]/comp_dict['Total'][1]],
        'DMCR': [comp_dict['DMCR'][0]/comp_dict['Total'][0],comp_dict['DMCR'][1]/comp_dict['Total'][1]],
        'UFER': [comp_dict['UFER'][0]/comp_dict['Total'][0],comp_dict['UFER'][1]/comp_dict['Total'][1]],
    }
    estilos_cargas.comparativo_style(grupo=grupo,comp_dict=comp_dict,writer=writer,pct_dict=pct_dict)
#--------------------------------------------------------------------------------------------------------