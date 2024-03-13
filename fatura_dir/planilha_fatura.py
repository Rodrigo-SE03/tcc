import pandas as pd
import copy
import os
from fatura_dir import estilos_fatura
import math

#Função geral de criação de planilhas
def criar_planilha(fatura_dict,nome,folder):
    fatura_dict = copy.deepcopy(fatura_dict)
    writer = pd.ExcelWriter(f'{folder}/{nome}',engine="xlsxwriter")
    tab_geral(fatura_dict=fatura_dict,writer=writer)
    tab_analise(fatura_dict=fatura_dict,writer=writer)
    writer.close()
#--------------------------------------------------------------------------------------------------------
    
#Criação da aba com as informações atuais de demanda e consumo da unidade consumidora
def tab_geral(fatura_dict,writer):

    if 'Demanda Verde Ultrapassada (atual)' in fatura_dict['Demanda'].keys():
        dados_dict = {
            'Mês': fatura_dict['Mês'],
            'Demanda Registrada HP': fatura_dict['Demanda']['Demanda P Medida'],
            'Demanda Registrada HFP': fatura_dict['Demanda']['Demanda FP Medida'],
            'Custo Demanda': fatura_dict['Demanda']['Custos com Demanda - Demanda Verde (atual)'],
            'Ultrapassagem Registrada': fatura_dict['Demanda']['Demanda Verde Ultrapassada (atual)'],
            'Custo Ultrapassagem': fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda Verde (atual)'],
            'Consumo HP': fatura_dict['Consumo']['Consumo P'],
            'Custo Consumo HP': fatura_dict['Consumo']['Custo Consumo - P Verde'],
            'Consumo HFP': fatura_dict['Consumo']['Consumo FP'],
            'Custo Consumo HFP': fatura_dict['Consumo']['Custo Consumo - FP Verde'],
            'Energia Reativa': fatura_dict['Reativo']['Consumo Reativo Medido'],
            'Custo Energia Reativa': fatura_dict['Reativo']['Custo do Consumo Reativo'],
            'Demanda Reativa': fatura_dict['Reativo']['Demanda Reativa Medida'],
            'Custo Demanda Reativa': fatura_dict['Reativo']['Custo da Demanda Reativa'],
        }
    else:
        dados_dict = {
            'Mês': fatura_dict['Mês'],
            'Demanda Registrada HP': fatura_dict['Demanda']['Demanda P Medida'],
            'Custo Demanda HP': fatura_dict['Demanda']['Custos com Demanda - Demanda P (atual)'],
            'Demanda Registrada HFP': fatura_dict['Demanda']['Demanda FP Medida'],
            'Custo Demanda HFP': fatura_dict['Demanda']['Custos com Demanda - Demanda FP (atual)'],
            'Ultrapassagem Registrada HP': fatura_dict['Demanda']['Demanda P Ultrapassada (atual)'],
            'Custo Ultrapassagem HP': fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P (atual)'],
            'Ultrapassagem Registrada HFP': fatura_dict['Demanda']['Demanda FP Ultrapassada (atual)'],
            'Custo Ultrapassagem HFP': fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP (atual)'],
            'Consumo HP': fatura_dict['Consumo']['Consumo P'],
            'Custo Consumo HP': fatura_dict['Consumo']['Custo Consumo - P Azul'],
            'Consumo HFP': fatura_dict['Consumo']['Consumo FP'],
            'Custo Consumo HFP': fatura_dict['Consumo']['Custo Consumo - FP Azul'],
            'Energia Reativa': fatura_dict['Reativo']['Consumo Reativo Medido'],
            'Custo Energia Reativa': fatura_dict['Reativo']['Custo do Consumo Reativo'],
            'Demanda Reativa': fatura_dict['Reativo']['Demanda Reativa Medida'],
            'Custo Demanda Reativa': fatura_dict['Reativo']['Custo da Demanda Reativa'],
        }
    df_dados = pd.DataFrame(dados_dict)
    df_dados.to_excel(writer, sheet_name="Geral", startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Geral"]
    (max_row, max_col) = df_dados.shape
    column_settings = [{"header": column} for column in df_dados.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    worksheet.autofit()
#--------------------------------------------------------------------------------------------------------
    
def tab_analise(fatura_dict,writer):
    print(list(reversed(fatura_dict['Mês'])))
    total_v = []
    total_a = []
    i=0
    while i < len(fatura_dict['Demanda']['Demanda Verde Medida']):
        total_v.append(fatura_dict['Demanda']['Custos com Demanda - Demanda Verde'][i]+fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda Verde'][i]+fatura_dict['Consumo']['Custo Consumo - P Verde'][i]+fatura_dict['Consumo']['Custo Consumo - FP Verde'][i])
        total_a.append(fatura_dict['Demanda']['Custos com Demanda - Demanda P'][i]+fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P'][i]+fatura_dict['Demanda']['Custos com Demanda - Demanda FP'][i]+fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP'][i]+fatura_dict['Consumo']['Custo Consumo - P Azul'][i]+fatura_dict['Consumo']['Custo Consumo - FP Azul'][i])
        i+=1

    verde_dict = {
        'Mês': list(reversed(fatura_dict['Mês'])),
        'Demanda': list(reversed(fatura_dict['Demanda']['Demanda Verde Medida'])),
        'Demanda - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda Verde'])),
        'Ultrapassagem - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda Verde'])),
        'Consumo HP': list(reversed(fatura_dict['Consumo']['Consumo P'])),
        'Consumo HP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Verde'])),
        'Consumo HFP': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
        'Consumo HFP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Verde'])),
        'Total': list(reversed(total_v))
    }

    azul_dict = {
        'Mês': list(reversed(fatura_dict['Mês'])),
        'Demanda HP': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
        'Demanda HP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda P'])),
        'Ultrapassagem HP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P'])),
        'Demanda HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
        'Demanda HFP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda FP'])),
        'Ultrapassagem HFP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP'])),
        'Consumo HP': list(reversed(fatura_dict['Consumo']['Consumo P'])),
        'Consumo HP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Azul'])),
        'Consumo HFP': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
        'Consumo HFP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Azul'])),
        'Total': list(reversed(total_a))
    }

    df_verde = pd.DataFrame(verde_dict)
    df_azul = pd.DataFrame(azul_dict)

    (max_row, max_col_v) = df_verde.shape
    (max_row, max_col_a) = df_azul.shape
    df_verde.to_excel(writer, sheet_name="Análise", startrow=1, header=False, index=False)
    df_azul.to_excel(writer, sheet_name="Análise", startrow=1,startcol=max_col_v+1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Análise"]
    worksheet.set_column(0, max_col_v+max_col_a - 1, 12)
    worksheet.autofit()