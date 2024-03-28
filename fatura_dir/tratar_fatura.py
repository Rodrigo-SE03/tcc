from PyPDF2 import PdfReader
import os

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

#Função para leitura dos dados da fatura
def ler_fatura(file,folder,tarifas,dem_c):

    reader = PdfReader(f'{folder}/{file}')
    page = reader.pages[1]
    # print(reader.pages[0].extract_text())
    text = page.extract_text()
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
    if isinstance(dem_c,float):
        ult_atual = []
        custos_ult_atual = []
        custos_dem_atual = []
        demandas = []
        i=0
        while i<len(dem_p):
            demandas.append(dem_fp[i] if dem_fp[i] > dem_p[i] else dem_p[i])
            ult_atual.append((demandas[i] - dem_c) if demandas[i]>dem_c*1.05 else 0)
            custos_ult_atual.append(ult_atual[i]*2*t_fp)
            custos_dem_atual.append(demandas[i]*t_fp)
            i+=1
        
    else:
        categoria = 'Azul'
        dem_c_fp = dem_c[0]
        dem_c_p = dem_c[1]

        ult_atual_fp = []
        custos_ult_fp_atual = []
        custos_dem_fp_atual = []
        i=0
        for dem in dem_fp:
            ult_atual_fp.append((dem-dem_c_fp) if dem>dem_c_fp*1.05 else 0)
            custos_ult_fp_atual.append(ult_atual_fp[i]*2*t_fp)
            custos_dem_fp_atual.append(dem*t_fp)
            i+=1
        
        ult_atual_p = []
        custos_ult_p_atual = []
        custos_dem_p_atual = []
        i=0
        for dem in dem_p:
            ult_atual_p.append((dem-dem_c_p) if dem>dem_c_p*1.05 else 0)
            custos_ult_p_atual.append(ult_atual_p[i]*2*t_p)
            custos_dem_p_atual.append(dem*t_p)
            i+=1
    

    min_dem_fp = 0
    max_dem_fp = int(max(dem_fp)/5)*5+500

    min_dem_p = int(min(dem_p)/5)*5
    max_dem_p = int(max(dem_p)/5)*5+300

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
            'Demanda FP Medida': dem_fp,
            'Demanda FP Ultrapassada': ult_fp,
            'Demanda P Medida': dem_p,
            'Demanda P Ultrapassada': ult_p,
            'Custos com Demanda - Demanda Verde': custos_dem_v,
            'Custos com Demanda - Demanda Verde (atual)': custos_dem_atual,               
            'Custos com Ultrapassagem - Demanda Verde': custos_ult_v,
            'Custos com Ultrapassagem - Demanda Verde (atual)': custos_ult_atual,
            'Custos com Demanda - Demanda FP': custos_dem_fp,             
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
            'Demanda P Medida': dem_p,
            'Demanda P Ultrapassada': ult_p,
            'Demanda P Ultrapassada (atual)': ult_atual_p,
            'Custos com Demanda - Demanda Verde': custos_dem_v,             
            'Custos com Ultrapassagem - Demanda Verde': custos_ult_v,
            'Custos com Demanda - Demanda FP': custos_dem_fp, 
            'Custos com Demanda - Demanda FP (atual)': custos_dem_fp_atual,              
            'Custos com Ultrapassagem - Demanda FP': custos_ult_fp,
            'Custos com Ultrapassagem - Demanda FP (atual)': custos_ult_fp_atual,
            'Custos com Demanda - Demanda P': custos_dem_p,
            'Custos com Demanda - Demanda P (atual)': custos_dem_p_atual,
            'Custos com Ultrapassagem - Demanda P': custos_ult_p,
            'Custos com Ultrapassagem - Demanda P (atual)': custos_ult_p_atual
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
    print(hr_con)
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


#Função para verificar se os dados foram inseridos corretamente
def verificar_save(fatura_dict):
    if len(fatura_dict) == 0:
        return 'Carregue uma fatura'
    return 'Arquivo salvo com sucesso'
#--------------------------------------------------------------------------------------------------------