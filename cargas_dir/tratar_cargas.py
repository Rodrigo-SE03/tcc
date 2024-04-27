import pandas as pd
import os

#Função para inserir uma carga na lista de elementos ou atualizar alguma já inserida
def nova_carga(cargas,form):
    if form.nome_equip.data in cargas['Carga']:
        id = cargas['Carga'].index(form.nome_equip.data)

        cargas['Carga'][id]=form.nome_equip.data
        cargas['Potência (kW)'][id]=form.potencia.data
        cargas['FP'][id]=form.fp.data
        cargas['FP - Tipo'][id]=form.fp_tipo.data
        cargas['Quantidade'][id]=form.qtd.data
        cargas['Início'][id]=form.hr_inicio.data
        cargas['Fim'][id]=form.hr_fim.data
        cargas['Remover'][id]='Remover'
    else:
        cargas['Carga'].append(form.nome_equip.data)
        cargas['Potência (kW)'].append(form.potencia.data)
        cargas['FP'].append(form.fp.data)
        cargas['FP - Tipo'].append(form.fp_tipo.data)
        cargas['Quantidade'].append(form.qtd.data)
        cargas['Início'].append(form.hr_inicio.data)
        cargas['Fim'].append(form.hr_fim.data)
        cargas['Remover'].append('Remover')
#--------------------------------------------------------------------------------------------------------

#Função para remover uma carga da lista de elementos
def remover_carga(cargas,i):
    nome = ''
    for keys in cargas:
        if keys == 'Carga':
            nome = cargas[keys][i]
        cargas[keys].pop(i)
    return nome
#--------------------------------------------------------------------------------------------------------

#Função para carregar a planilha pré construída com dados de cargas
def carregar_cargas(file,folder): 
    cargas = {
        'Carga':[],
        'Potência (kW)':[],
        'FP':[],
        'FP - Tipo':[],
        'Quantidade':[],
        'Início':[],
        'Fim':[],
        'Remover': []
    }
    df = pd.read_excel(f'{folder}/{file}')
    os.remove(f'{folder}/{file}')
    print(df)
    try:
        if df.iloc[0,7] == 'Validar':
            pass
        else:
            return 'Arquivo inválido'
    except:
        return 'Arquivo inválido'
    for col in df.columns:
        for item in df[col]:
            if col in cargas.keys():
                cargas[col].append(item)
    i=0
    while i < len(cargas['Carga']):
        cargas['Remover'].append('Remover')
        i+=1
    return cargas
#--------------------------------------------------------------------------------------------------------

#Função para verificar se os dados foram inseridos corretamente
def verificar_save(cargas_dict,tarifas_dict,h_p,dias):
    if len(cargas_dict['Carga']) == 0:
        return 'Insira pelo menos uma carga'
    
    if not isinstance(h_p,float) or h_p == 0:
        return 'Preencha corretamente os dados de horário de ponta e dias úteis'
    if not isinstance(dias,int) or dias == 0:
        return 'Preencha corretamente os dados de horário de ponta e dias úteis'

    for cat in tarifas_dict.keys():
        if cat == 'convencional':
            if tarifas_dict[cat] == 0: 
                return 'Informe as tarifas praticadas'
        else:
            if cat == 'te':
                if tarifas_dict[cat] != 0: continue
            for t in tarifas_dict[cat]:
                if t  == 0:
                    return 'Informe as tarifas praticadas'
    
    return 'Arquivo salvo com sucesso'
#--------------------------------------------------------------------------------------------------------