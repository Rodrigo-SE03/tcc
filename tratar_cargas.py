import pandas as pd
import os

#Função para inserir uma carga na lista de elementos ou atualizar alguma já inserida
def nova_carga(cargas,form):
    if form.nome_equip.data in cargas['Carga']:
        id = cargas['Carga'].index(form.nome_equip.data)

        cargas['Carga'][id]=form.nome_equip.data
        cargas['Potência'][id]=form.potencia.data
        cargas['FP'][id]=form.fp.data
        cargas['FP - Tipo'][id]=form.fp_tipo.data
        cargas['Quantidade'][id]=form.qtd.data
        cargas['Início'][id]=form.hr_inicio.data
        cargas['Fim'][id]=form.hr_fim.data
        cargas['Remover'][id]='Remover'
    else:
        cargas['Carga'].append(form.nome_equip.data)
        cargas['Potência'].append(form.potencia.data)
        cargas['FP'].append(form.fp.data)
        cargas['FP - Tipo'].append(form.fp_tipo.data)
        cargas['Quantidade'].append(form.qtd.data)
        cargas['Início'].append(form.hr_inicio.data)
        cargas['Fim'].append(form.hr_fim.data)
        cargas['Remover'].append('Remover')
#--------------------------------------------------------------------------------------------------------

#Função para remover uma carga da lista de elementos
def remover_carga(cargas,i):
    for keys in cargas:
        cargas[keys].pop(i)
#--------------------------------------------------------------------------------------------------------

#Função para carregar a planilha pré construída com dados de cargas
def carregar_cargas(file,folder): 
    cargas = {
        'Carga':[],
        'Potência':[],
        'FP':[],
        'FP - Tipo':[],
        'Quantidade':[],
        'Início':[],
        'Fim':[],
        'Remover': []
    }
    df = pd.read_excel(f'{folder}/{file}')
    os.remove(f'{os.getcwd()}/{folder}/{file}')
    for col in df.columns:
        for item in df[col]:
            cargas[col].append(item)
    i=0
    while i < len(cargas['Carga']):
        cargas['Remover'].append('Remover')
        i+=1
    return cargas
#--------------------------------------------------------------------------------------------------------

#Função para verificar se os dados foram inseridos corretamente
def verificar_save(cargas_dict,tarifas_dict):
    if len(cargas_dict['Carga']) == 0:
        return 'Insira pelo menos uma carga'

    for cat in tarifas_dict.keys():
        if cat == 'convencional':
            if tarifas_dict[cat] == 0: 
                return 'Informe as tarifas praticadas'
        else:
            for t in tarifas_dict[cat]:
                if t  == 0:
                    return 'Informe as tarifas praticadas'
    
    return 'Arquivo salvo com sucesso'
    #--------------------------------------------------------------------------------------------------------