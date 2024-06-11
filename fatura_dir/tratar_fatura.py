from PyPDF2 import PdfReader
import os
from datetime import date, timedelta
import pandas as pd
import re

def demanda_contratada(form_fatura):
    dem_c = 0
    if form_fatura.dem_c_p.data == 0:
        dem_c = form_fatura.dem_c_fp.data
    else:
        dem_c = [form_fatura.dem_c_fp.data,form_fatura.dem_c_p.data]
    return dem_c

#Função para converter o texto extraído para valor numérico
def converter(valor):
    valor = valor.replace('.','')
    valor = valor.replace(',','.')
    return float(valor)
#--------------------------------------------------------------------------------------------------------

def get_mes_num(mes):
    if mes == 'JAN': mes = 1
    elif mes == 'FEV': mes = 2
    elif mes == 'MAR': mes = 3
    elif mes == 'ABR': mes = 4
    elif mes == 'MAI': mes = 5
    elif mes == 'JUN': mes = 6
    elif mes == 'JUL': mes = 7
    elif mes == 'AGO': mes = 8
    elif mes == 'SET': mes = 9
    elif mes == 'OUT': mes = 10
    elif mes == 'NOV': mes = 11
    elif mes == 'DEZ': mes = 12
    return mes

def get_mes_nom(mes):
    if mes == 1: mes = 'JAN'
    elif mes == 2: mes = 'FEV'
    elif mes == 3: mes = 'MAR'
    elif mes == 4: mes = 'ABR'
    elif mes == 5: mes = 'MAI'
    elif mes == 6: mes = 'JUN'
    elif mes == 7: mes = 'JUL'
    elif mes == 8: mes = 'AGO'
    elif mes == 9: mes = 'SET'
    elif mes == 10: mes = 'OUT'
    elif mes == 11: mes = 'NOV'
    elif mes == 12: mes = 'DEZ'
    return mes

#Função para definir os meses da análise
def definir_meses(data):
    meses = []
    anos = []
    new_m = date(data[1],get_mes_num(data[0]),1)
    meses.append(get_mes_nom(new_m.month))
    anos.append(new_m.year)
    for i in range(0,11):
        new_m = (new_m - timedelta(days=1)).replace(day=1)
        meses.append(get_mes_nom(new_m.month))
        anos.append(new_m.year)
    
    return [meses,anos]
#--------------------------------------------------------------------------------------------------------

#Função para salvar os dados no modelo manual
def dados_manual(form_manual,dem_c,tarifas,meses,anos):

    dem_p = form_manual.demanda_p.data
    dem_fp = form_manual.demanda_fp.data
    dem_r = form_manual.dmcr.data
    con_p = form_manual.consumo_p.data
    con_fp = form_manual.consumo_fp.data
    con_r = form_manual.ufer.data
    hr_con = form_manual.consumo_hr.data
    hr_r = form_manual.ufer_hr.data

    historico_dict={
        'mes':meses,
        'ano': anos,
        'demanda_p':dem_p,
        'demanda_fp':dem_fp,
        'dmcr':dem_r,
        'consumo_p': con_p,
        'consumo_fp': con_fp,
        'consumo_hr': hr_con,
        'ufer': con_r,
        'ufer_hr': hr_r
    }

    data = [meses[0],anos[0]]
    meses_formatado = []
    for i in range(0,12):
        meses_formatado.append(f'{meses[i]}/{str(anos[i])[-2:]}')

    demandas_dict = calc_demanda(dem_p=dem_p,dem_fp=dem_fp,dem_c=dem_c,mes=meses_formatado,tarifas=tarifas)
    consumos_dict = calc_consumos(con_fp=con_fp,con_p=con_p,hr_con=hr_con,mes=meses_formatado,tarifas=tarifas)
    reativos_dict = calc_reativos(con_r=con_r,dem_r=dem_r,hr_r=hr_r,mes=meses_formatado,tarifas=tarifas)

    fatura_dict = {
        'Mês': meses_formatado,
        'Demanda': demandas_dict,
        'Consumo': consumos_dict,
        'Reativo': reativos_dict
    }
    return [fatura_dict,historico_dict,data]

#--------------------------------------------------------------------------------------------------------

#Função para leitura dos dados de histórico das faturas no modelo em excel
def ler_excel(file,folder,tarifas,dem_c):

    df = pd.read_excel(f'{folder}/{file}')
    os.remove(f'{folder}/{file}')
    try:
        if df.iloc[0,9] == 'Validar':
            pass
        else:
            print('else')
            return 'Arquivo inválido'
    except:
        print('except')
        return 'Arquivo inválido'
    
    m0 = df['Mês'].tolist()[0]
    data = [0,0]
    data[0] = m0.split('/')[0]
    data[1] = int(f'20{m0.split('/')[1]}')
    meses,anos = definir_meses(data)
    dem_p = df['Demanda Registrada na Ponta'].tolist()
    dem_fp = df['Demanda Registrada Fora Ponta'].tolist()
    dem_r = df['DMCR'].tolist()
    con_p = df['Consumo na Ponta'].tolist()
    con_fp = df['Consumo Fora Ponta'].tolist()
    con_r = df['UFER'].tolist()
    hr_con = df['Consumo no Horário Reservado'].tolist()
    hr_r = df['UFER Horário Reservado'].tolist()

    meses_formatado = []
    for i in range(0,12):
        meses_formatado.append(f'{meses[i]}/{str(anos[i])[-2:]}')

    demandas_dict = calc_demanda(dem_p=dem_p,dem_fp=dem_fp,dem_c=dem_c,mes=meses_formatado,tarifas=tarifas)
    consumos_dict = calc_consumos(con_fp=con_fp,con_p=con_p,hr_con=hr_con,mes=meses_formatado,tarifas=tarifas)
    reativos_dict = calc_reativos(con_r=con_r,dem_r=dem_r,hr_r=hr_r,mes=meses_formatado,tarifas=tarifas)

    fatura_dict = {
        'Mês': meses_formatado,
        'Demanda': demandas_dict,
        'Consumo': consumos_dict,
        'Reativo': reativos_dict
    }

    historico_dict={
        'mes':meses,
        'ano': anos,
        'demanda_p':dem_p,
        'demanda_fp':dem_fp,
        'dmcr':dem_r,
        'consumo_p': con_p,
        'consumo_fp': con_fp,
        'consumo_hr': hr_con,
        'ufer': con_r,
        'ufer_hr': hr_r
    }

    return [fatura_dict,historico_dict,data]
            
#--------------------------------------------------------------------------------------------------------

#Função para leitura dos dados da fatura
def ler_fatura(file,folder,tarifas,dem_c):

    reader = PdfReader(f'{folder}/{file}')
    page = reader.pages[1]
    if 'Motivo' in page.extract_text(): 
        page = reader.pages[0]
        text = page.extract_text()
        text = re.findall(r'( ([DNOSAJMF][A-Z]+ [ \/0-9,A-Z]+\n){13})',text)[0][0]
        text = text[1:]
    else:
        text = page.extract_text()
    if 'EQUATORIAL' not in reader.pages[0].extract_text():
        return 'Arquivo inválido'
    
    text = text.replace('\n',' ')
    values = text.split(' ')
    # print(values)
    i=0
    while i<4:
        values.pop()
        i+=1
    
    mes = ['']
    dem_p = []
    dem_fp = []
    dem_r = []
    con_p = []
    con_fp = []
    con_r = []
    hr_con = []
    hr_r = []
    i=0
    j=0
    while j < 12:
        while i < 11:
            if i<=2:
                mes[j] = ''.join([mes[j],values[i+j*11]])
            else:
                if i == 3:
                    con_p.append(converter(values[i+j*11]))
                elif i == 4:
                    con_fp.append(converter(values[i+j*11]))
                elif i == 5:
                    dem_p.append(converter(values[i+j*11]))
                elif i == 6:
                    dem_fp.append(converter(values[i+j*11]))
                elif i == 7:
                    con_r.append(converter(values[i+j*11]))
                elif i == 8:
                    dem_r.append(converter(values[i+j*11]))
                elif i == 9:
                    hr_con.append(converter(values[i+j*11]))
                elif i == 10:
                    hr_r.append(converter(values[i+j*11]))
            i+=1
        i=0
        mes.append('')
        j+=1
    mes.pop()

    demandas_dict = calc_demanda(dem_p=dem_p,dem_fp=dem_fp,dem_c=dem_c,mes=mes,tarifas=tarifas)
    consumos_dict = calc_consumos(con_fp=con_fp,con_p=con_p,hr_con=hr_con,mes=mes,tarifas=tarifas)
    reativos_dict = calc_reativos(con_r=con_r,dem_r=dem_r,hr_r=hr_r,mes=mes,tarifas=tarifas)

    fatura_dict = {
        'Mês': mes,
        'Demanda': demandas_dict,
        'Consumo': consumos_dict,
        'Reativo': reativos_dict
    }

    os.remove(f'{folder}/{file}')
    return fatura_dict
#--------------------------------------------------------------------------------------------------------
    
#Função para o cálculo da demanda contratada ideal
def calc_demanda(dem_p,dem_fp,dem_c,mes,tarifas):
    t_fp = tarifas['verde'][2]
    t_p = tarifas['azul'][3]
    categoria = 'Verde'
    if isinstance(dem_c,list) == False:
        ult_atual = []
        custos_ult_atual = []
        custos_dem_atual = []
        dem_nao_ut = []
        custos_dem_nao_ut = []
        demandas = []
        i=0
        while i<len(dem_p):
            demandas.append(dem_fp[i] if dem_fp[i] > dem_p[i] else dem_p[i])
            dem_nao_ut.append((dem_c-demandas[i]) if dem_c>demandas[i] else 0)
            ult_atual.append((demandas[i] - dem_c) if demandas[i]>dem_c*1.05 else 0)
            custos_ult_atual.append(ult_atual[i]*2*t_fp)
            custos_dem_atual.append(t_fp*demandas[i])
            custos_dem_nao_ut.append(t_fp*((dem_c-demandas[i]) if dem_c>demandas[i] else 0))
            i+=1
        
    else:
        categoria = 'Azul'
        dem_c_fp = dem_c[0]
        dem_c_p = dem_c[1]

        ult_atual_fp = []
        dem_nao_ut_fp = []
        custos_ult_fp_atual = []
        custos_dem_fp_atual = []
        custos_dem_nao_ut_fp = []
        i=0
        for dem in dem_fp:
            ult_atual_fp.append((dem-dem_c_fp) if dem>dem_c_fp*1.05 else 0)
            dem_nao_ut_fp.append((dem_c_fp-dem_fp[i]) if dem_c_fp>dem_fp[i] else 0)
            custos_ult_fp_atual.append(ult_atual_fp[i]*2*t_fp)
            custos_dem_fp_atual.append(t_fp*dem_fp[i])
            custos_dem_nao_ut_fp.append(t_fp*((dem_c_fp-dem_fp[i]) if dem_c_fp>dem_fp[i] else 0))
            i+=1
        
        ult_atual_p = []
        dem_nao_ut_p = []
        custos_ult_p_atual = []
        custos_dem_p_atual = []
        custos_dem_nao_ut_p = []
        i=0
        for dem in dem_p:
            ult_atual_p.append((dem-dem_c_p) if dem>dem_c_p*1.05 else 0)
            dem_nao_ut_p.append((dem_c_p-dem_p[i]) if dem_c_p>dem_p[i] else 0)
            custos_ult_p_atual.append(ult_atual_p[i]*2*t_p)
            custos_dem_p_atual.append(t_p*(dem_c_p if dem_c_p>dem_p[i] else dem_p[i]))
            custos_dem_nao_ut_p.append(t_p*((dem_c_p-dem_p[i]) if dem_c_p>dem_p[i] else 0))
            i+=1
    

    min_dem_fp = 30
    max_dem_fp = int(max(dem_fp)/5)*5+500

    min_dem_p = 30
    max_dem_p = int(max(dem_p)/5)*5+500

    demandas = dem_fp

    #Cálculo da demanda ideal fora ponta
    custos = []
    custos_dem_fp = []
    custos_ult_fp = []
    ult_fp = []
    dem_c_v_list = []
    for new_dem_c in range(min_dem_fp,max_dem_fp,5):
        new_ult_fp = []
        new_custo_fp = []
        new_custo_dem = []
        new_custo_ult = []
        dem_c_v_list.append(new_dem_c)
        i=0
        for dem in demandas:
            if dem >= 1.05*new_dem_c:
                new_ult_fp.append(dem-new_dem_c)
            else:
                new_ult_fp.append(0)
            new_custo_ult.append(new_ult_fp[i]*2*t_fp)
            new_custo_dem.append(t_fp*(demandas[i] if demandas[i] > new_dem_c else new_dem_c ))
            new_custo_fp.append(new_custo_ult[i]+new_custo_dem[i])
            i+=1
        ult_fp.append(new_ult_fp)
        custos_ult_fp.append(new_custo_ult)
        custos_dem_fp.append(new_custo_dem)
        custos.append(new_custo_fp)
    
    i=0
    id = 0
    custo_v = []
    minimo = 999999999999
    while i<len(custos):
        custo_v.append(sum(custos[i]))
        if minimo > sum(custos[i]):
            minimo = sum(custos[i])
            id = i
        i+=1

    new_dem_c_fp = dem_c_v_list[id]
    ult_fp = ult_fp[id]
    custos_ult_fp = custos_ult_fp[id]
    custos_dem_fp = custos_dem_fp[id]
    #--------------------------------------------------------------------------------------------------------

    #Cálculo da demanda ideal para modalidade verde
    demandas = []
    i=0
    while i<len(dem_p):
        demandas.append(dem_fp[i] if dem_fp[i] > dem_p[i] else dem_p[i])
        i+=1
    dem_v = demandas
    custos = []
    custos_dem_v = []
    custos_ult_v = []
    ult_v = []
    dem_c_fp_list = []
    for new_dem_c in range(min_dem_fp,max_dem_fp,5):
        new_ult_fp = []
        new_custo_fp = []
        new_custo_dem = []
        new_custo_ult = []
        dem_c_fp_list.append(new_dem_c)
        i=0
        for dem in demandas:
            if dem >= 1.05*new_dem_c:
                new_ult_fp.append(dem-new_dem_c)
            else:
                new_ult_fp.append(0)
            new_custo_ult.append(new_ult_fp[i]*2*t_fp)
            new_custo_dem.append(t_fp*(demandas[i] if demandas[i] > new_dem_c else new_dem_c ))
            new_custo_fp.append(new_custo_ult[i]+new_custo_dem[i])
            i+=1
        ult_v.append(new_ult_fp)
        custos_ult_v.append(new_custo_ult)
        custos_dem_v.append(new_custo_dem)
        custos.append(new_custo_fp)
    
    i=0
    id = 0
    minimo = 999999999999
    while i<len(custos):
        if minimo > sum(custos[i]):
            minimo = sum(custos[i])
            id = i
        i+=1

    new_dem_c_fp = dem_c_fp_list[id]
    custo_fp = custos[id]
    ult_v = ult_v[id]
    custos_ult_v = custos_ult_v[id]
    custos_dem_v = custos_dem_v[id]
    #--------------------------------------------------------------------------------------------------------
        
    #Cálculo da demanda ideal na ponta
    demandas = dem_p
    custos = []
    custos_dem_p = []
    custos_ult_p = []
    ult_p = []
    dem_c_p_list = []
    for new_dem_c in range(min_dem_p,max_dem_p,5):
        new_ult_p = []
        new_custo_p = []
        new_custo_dem = []
        new_custo_ult = []
        dem_c_p_list.append(new_dem_c)
        i=0
        for dem in demandas:
            if dem >= 1.05*new_dem_c:
                new_ult_p.append(dem-new_dem_c)
            else:
                new_ult_p.append(0)

            new_custo_ult.append(new_ult_p[i]*2*t_p)
            new_custo_dem.append(t_p*(demandas[i] if demandas[i] > new_dem_c else new_dem_c ))
            new_custo_p.append(new_custo_ult[i]+new_custo_dem[i])
            i+=1
        ult_p.append(new_ult_p)
        custos_ult_p.append(new_custo_ult)
        custos_dem_p.append(new_custo_dem)
        custos.append(new_custo_p)
    
    i=0
    id = 0
    minimo = 999999999999
    while i<len(custos):
        if minimo > sum(custos[i]):
            minimo = sum(custos[i])
            id = i
        i+=1

    new_dem_c_p = dem_c_p_list[id]
    custo_p = custos[id]
    ult_p = ult_p[id]
    custos_ult_p = custos_ult_p[id]
    custos_dem_p = custos_dem_p[id]
    #--------------------------------------------------------------------------------------------------------

    if categoria == 'Verde':
        demandas_dict = {
            'Demanda Contratada FP Indicada': new_dem_c_fp,
            'Demanda Contratada P Indicada': new_dem_c_p,
            'Demanda Contratada Atual': dem_c,
            'Demanda Verde Medida': dem_v,
            'Demanda Verde Ultrapassada (atual)': ult_atual,
            'Demanda Verde Ultrapassada': ult_v,
            'Demanda Verde Não Utilizada': dem_nao_ut,
            'Demanda FP Medida': dem_fp,
            'Demanda FP Ultrapassada': ult_fp,
            'Demanda P Medida': dem_p,
            'Demanda P Ultrapassada': ult_p,
            'Custos com Demanda - Demanda Verde': custos_dem_v,
            'Custos com Demanda - Demanda Verde (atual)': custos_dem_atual,               
            'Custos com Ultrapassagem - Demanda Verde': custos_ult_v,
            'Custos com Ultrapassagem - Demanda Verde (atual)': custos_ult_atual,
            'Custos com Demanda - Demanda FP': custos_dem_fp,
            'Custos com Demanda Verde Não Utilizada': custos_dem_nao_ut,             
            'Custos com Ultrapassagem - Demanda FP': custos_ult_fp,
            'Custos com Demanda - Demanda P': custos_dem_p,
            'Custos com Ultrapassagem - Demanda P': custos_ult_p,
            'Lista de demandas contratadas': dem_c_v_list,
            'Lista de custos anuais por demanda contratada': custo_v
        }
    else:
        demandas_dict = {
            'Demanda Contratada FP Indicada': new_dem_c_fp,
            'Demanda Contratada P Indicada': new_dem_c_p,
            'Demanda Contratada FP Atual': dem_c[0],
            'Demanda Contratada P Atual': dem_c[1],
            'Demanda Verde Medida': dem_v,
            'Demanda Verde Ultrapassada': ult_v,
            'Demanda FP Medida': dem_fp,
            'Demanda FP Ultrapassada': ult_fp,
            'Demanda FP Ultrapassada (atual)': ult_atual_fp,
            'Demanda FP Não Utilizada': dem_nao_ut_fp,
            'Demanda P Medida': dem_p,
            'Demanda P Ultrapassada': ult_p,
            'Demanda P Ultrapassada (atual)': ult_atual_p,
            'Demanda P Não Utilizada': dem_nao_ut_p,
            'Custos com Demanda - Demanda Verde': custos_dem_v,             
            'Custos com Ultrapassagem - Demanda Verde': custos_ult_v,
            'Custos com Demanda - Demanda FP': custos_dem_fp, 
            'Custos com Demanda FP Não Utilizada': custos_dem_nao_ut_fp, 
            'Custos com Demanda - Demanda FP (atual)': custos_dem_fp_atual,              
            'Custos com Ultrapassagem - Demanda FP': custos_ult_fp,
            'Custos com Ultrapassagem - Demanda FP (atual)': custos_ult_fp_atual,
            'Custos com Demanda - Demanda P': custos_dem_p,
            'Custos com Demanda P Não Utilizada': custos_dem_nao_ut_p, 
            'Custos com Demanda - Demanda P (atual)': custos_dem_p_atual,
            'Custos com Ultrapassagem - Demanda P': custos_ult_p,
            'Custos com Ultrapassagem - Demanda P (atual)': custos_ult_p_atual,
            'Lista de demandas contratadas': dem_c_v_list,
            'Lista de custos anuais por demanda contratada': custo_v
        }
    # print(demandas_dict)
    return demandas_dict
#--------------------------------------------------------------------------------------------------------

#Função para cálculo dos custos dos consumos
def calc_consumos(con_fp,con_p,hr_con,tarifas,mes):
    t_fp_v = tarifas['verde'][0]
    t_p_v = tarifas['verde'][1]
    t_fp_a = tarifas['azul'][0]
    t_p_a = tarifas['azul'][1]
    # print(hr_con)
    consumo_fp = []
    custo_fp_v = []
    custo_p_v = []
    custo_fp_a = []
    custo_p_a = []
    i=0
    while i<len(con_fp):
        consumo_fp.append(con_fp[i]+hr_con[i])
        custo_fp_v.append((con_fp[i]+hr_con[i])*t_fp_v)
        custo_p_v.append(con_p[i]*t_p_v)
        custo_fp_a.append((con_fp[i]+hr_con[i])*t_fp_a)
        custo_p_a.append(con_p[i]*t_p_a)
        i+=1

    consumos_dict = {
        'Consumo FP':consumo_fp,
        'Consumo P':con_p,
        'Custo Consumo - FP Verde': custo_fp_v,
        'Custo Consumo - P Verde': custo_p_v,
        'Custo Consumo - FP Azul': custo_fp_a,
        'Custo Consumo - P Azul': custo_p_a
    }
    
    # print(consumos_dict)
    return consumos_dict
#--------------------------------------------------------------------------------------------------------

#Função para cálculo dos custos com energia reativa
def calc_reativos(dem_r,con_r,hr_r,tarifas,mes):
    t_r = tarifas['te']
    t_d = tarifas['verde'][2]
    custo_consumo_r = []
    custo_dem_r = []
    consumo_r = []
    i=0
    while i<len(dem_r):
        consumo_r.append(con_r[i]+hr_r[i])
        custo_consumo_r.append(consumo_r[i]*t_r)
        custo_dem_r.append(dem_r[i]*t_d)
        i+=1
    
    reativos_dict = {
        'Consumo Reativo Medido': consumo_r,
        'Demanda Reativa Medida': dem_r,
        'Custo do Consumo Reativo': custo_consumo_r,
        'Custo da Demanda Reativa': custo_dem_r
    }

    # print(reativos_dict)
    return reativos_dict
#--------------------------------------------------------------------------------------------------------

#Função para verificar se os dados foram inseridos corretamente
def verificar_save(fatura_dict):
    if len(fatura_dict) == 0:
        return 'Carregue uma fatura'
    return 'Arquivo salvo com sucesso'
#--------------------------------------------------------------------------------------------------------