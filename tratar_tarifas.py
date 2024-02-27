import pandas as pd
import os

def registrar_tarifas(tarifas,form,grupo):
    if grupo == 'Grupo B':
        tarifas = {'convencional':form.convencional.data,
                   'branca': [form.branca_fp.data,form.branca_i.data,form.branca_p.data]}
        
    elif grupo == 'Grupo A':
        tarifas = {'verde': [form.verde_fp.data,form.verde_p.data,form.verde_dem.data],
                   'azul': [form.azul_fp.data,form.azul_p.data,form.azul_dem_fp.data,form.azul_dem_p.data]}
    
    return tarifas

def carregar_tarifas(file,folder,grupo): 
    df = pd.read_excel(f'{folder}/{file}')
    if grupo == 'Grupo B':
        tarifas = {'convencional':df.iloc[2,3],
                   'branca': [df.iloc[6,3],df.iloc[7,3],df.iloc[8,3]]}
    elif grupo == 'Grupo A':
        tarifas = {'verde': [df.iloc[12,3],df.iloc[13,3],df.iloc[14,3]],
                   'azul': [df.iloc[18,3],df.iloc[19,3],df.iloc[20,3],df.iloc[21,3]]}
    os.remove(f'{os.getcwd()}/{folder}/{file}')
    
    return tarifas
    
    