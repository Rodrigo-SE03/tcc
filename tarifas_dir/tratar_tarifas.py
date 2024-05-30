import pandas as pd
import os

def registrar_tarifas(tarifas,form,grupo):
    if grupo == 'Grupo B':
        tarifas = {'convencional':form.convencional.data,
                   'branca': [form.branca_fp.data,form.branca_i.data,form.branca_p.data]}
        
    elif grupo == 'Grupo A':
        tarifas = {'verde': [form.verde_fp.data,form.verde_p.data,form.verde_dem.data],
                   'azul': [form.azul_fp.data,form.azul_p.data,form.azul_dem_fp.data,form.azul_dem_p.data],
                   'te': form.te.data}
    
    return tarifas

def carregar_tarifas(file,folder,grupo): 
    df = pd.read_excel(f'{folder}/{file}')
    try:
        if df.iloc[2,4] == 'Validar':
            pass
        else:
            return 'Arquivo inválido'
    except:
        return 'Arquivo inválido'
    if grupo == 'Grupo B':
        tarifas = {'convencional':df.iloc[2,3],
                   'branca': [df.iloc[6,3],df.iloc[7,3],df.iloc[8,3]]}
        if tarifas['convencional'] != tarifas['convencional']:
            return 'Arquivo inválido'
        elif tarifas['branca'][0] != tarifas['branca'][0] or tarifas['branca'][1] != tarifas['branca'][1] or tarifas['branca'][2] != tarifas['branca'][2]: 
            return 'Arquivo inválido'
        else:
            pass
    elif grupo == 'Grupo A':
        tarifas = {'verde': [df.iloc[12,3],df.iloc[13,3],df.iloc[14,3]],
                   'azul': [df.iloc[18,3],df.iloc[19,3],df.iloc[20,3],df.iloc[21,3]],
                   'te': df.iloc[25,3]}
        print(type(tarifas['te']))
        for key in tarifas.keys():
            if key != 'te':
                for i in range(0,len(tarifas[key])):
                    if tarifas[key][i] != tarifas[key][i]:
                        return 'Arquivo inválido'
            else:
                if tarifas[key] != tarifas[key]:
                    return 'Arquivo inválido'
        
    os.remove(f'{folder}/{file}')
    
    return tarifas