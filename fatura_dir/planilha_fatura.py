import pandas as pd
import copy
import os
from fatura_dir import estilos_fatura,graficos
import math

#Função geral de criação de planilhas
def criar_planilha(fatura_dict,nome,folder):
    categoria = 'Verde' if 'Demanda Verde Ultrapassada (atual)' in fatura_dict['Demanda'].keys() else 'Azul'
    fatura_dict = copy.deepcopy(fatura_dict)
    writer = pd.ExcelWriter(f'{folder}/{nome}',engine="xlsxwriter")
    tab_geral(fatura_dict=fatura_dict,writer=writer,categoria=categoria)
    tab_analise(fatura_dict=fatura_dict,writer=writer,categoria=categoria)
    tab_recomendado(fatura_dict=fatura_dict,writer=writer)
    writer.close()
#--------------------------------------------------------------------------------------------------------
    
#Criação da aba com as informações atuais de demanda e consumo da unidade consumidora
def tab_geral(fatura_dict,writer,categoria):

    if categoria == 'Verde':
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Demanda Registrada HP': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
            'Demanda Registrada HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
            'Custo Demanda': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda Verde (atual)'])),
            'Ultrapassagem Registrada': list(reversed(fatura_dict['Demanda']['Demanda Verde Ultrapassada (atual)'])),
            'Custo Ultrapassagem': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda Verde (atual)'])),
            'Consumo HP': list(reversed(fatura_dict['Consumo']['Consumo P'])),
            'Custo Consumo HP': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Verde'])),
            'Consumo HFP': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
            'Custo Consumo HFP': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Verde'])),
            'Energia Reativa': list(reversed(fatura_dict['Reativo']['Consumo Reativo Medido'])),
            'Custo Energia Reativa': list(reversed(fatura_dict['Reativo']['Custo do Consumo Reativo'])),
            'Demanda Reativa': list(reversed(fatura_dict['Reativo']['Demanda Reativa Medida'])),
            'Custo Demanda Reativa': list(reversed(fatura_dict['Reativo']['Custo da Demanda Reativa'])),
        }
    else:
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Demanda Registrada HP': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
            'Custo Demanda HP': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda P (atual)'])),
            'Demanda Registrada HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
            'Custo Demanda HFP': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda FP (atual)'])),
            'Ultrapassagem Registrada HP': list(reversed(fatura_dict['Demanda']['Demanda P Ultrapassada (atual)'])),
            'Custo Ultrapassagem HP': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P (atual)'])),
            'Ultrapassagem Registrada HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Ultrapassada (atual)'])),
            'Custo Ultrapassagem HFP': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP (atual)'])),
            'Consumo HP': list(reversed(fatura_dict['Consumo']['Consumo P'])),
            'Custo Consumo HP': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Azul'])),
            'Consumo HFP': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
            'Custo Consumo HFP': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Azul'])),
            'Energia Reativa': list(reversed(fatura_dict['Reativo']['Consumo Reativo Medido'])),
            'Custo Energia Reativa': list(reversed(fatura_dict['Reativo']['Custo do Consumo Reativo'])),
            'Demanda Reativa': list(reversed(fatura_dict['Reativo']['Demanda Reativa Medida'])),
            'Custo Demanda Reativa': list(reversed(fatura_dict['Reativo']['Custo da Demanda Reativa'])),
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
    graficos.graf_reativos(categoria=categoria,sheet_name="Geral",workbook=workbook,worksheet=worksheet)
#--------------------------------------------------------------------------------------------------------

#Criação da aba com a comparação entre as modalidades azul e verde para a unidade consumidora, considerando as demandas contratadas ideais calculadas
def tab_analise(fatura_dict,writer,categoria):
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
        'Ultrapassagem': list(reversed(fatura_dict['Demanda']['Demanda Verde Ultrapassada'])),
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
        'Ultrapassagem HP': list(reversed(fatura_dict['Demanda']['Demanda P Ultrapassada'])),
        'Ultrapassagem HP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P'])),
        'Demanda HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
        'Demanda HFP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda FP'])),
        'Ultrapassagem HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Ultrapassada'])),
        'Ultrapassagem HFP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP'])),
        'Consumo HP': list(reversed(fatura_dict['Consumo']['Consumo P'])),
        'Consumo HP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Azul'])),
        'Consumo HFP': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
        'Consumo HFP - Faturado': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Azul'])),
        'Total': list(reversed(total_a))
    }
    total = [sum(verde_dict['Total']),sum(azul_dict['Total'])]
    workbook = writer.book
    worksheet = workbook.add_worksheet('Comparativo')
    estilos_fatura.comparar_geral_style(workbook=workbook,worksheet=worksheet,verde_dict=verde_dict,azul_dict=azul_dict,fatura_dict=fatura_dict,categoria=categoria,total=total)
#--------------------------------------------------------------------------------------------------------

#Criação da aba com informações da estimativa de gastos e economia da unidade de acordo com a modalidade e demanda contratada indicada
def tab_recomendado(fatura_dict,writer):
    categoria = 0
    if 'Demanda Verde Ultrapassada (atual)' in fatura_dict['Demanda'].keys():
        dem_c = fatura_dict['Demanda']['Demanda Contratada Atual']
        dem_rec = fatura_dict['Demanda']['Demanda Contratada FP Indicada']
        dem_rec_list = []
        dem_c_list = []
        lim = []
        for m in fatura_dict['Mês']:
            dem_c_list.append(dem_c)
            dem_rec_list.append(dem_rec)
            lim.append(dem_rec*1.05)
        categoria = 'Verde'
    else:                                                                             
        dem_c_fp = fatura_dict['Demanda']['Demanda Contratada FP Atual']
        dem_c_p = fatura_dict['Demanda']['Demanda Contratada P Atual']
        dem_rec_fp = fatura_dict['Demanda']['Demanda Contratada FP Indicada']
        dem_rec_p = fatura_dict['Demanda']['Demanda Contratada P Indicada']
        dem_rec_fp_list = []
        dem_rec_p_list = []
        dem_c_fp_list = []
        dem_c_p_list = []
        lim_fp = []
        lim_p = []
        for m in fatura_dict['Mês']:
            dem_c_fp_list.append(dem_c_fp)
            dem_c_p_list.append(dem_c_p)
            dem_rec_fp_list.append(dem_rec_fp)
            dem_rec_p_list.append(dem_rec_p)
            lim_fp.append(dem_rec_fp*1.05)
            lim_p.append(dem_rec_p*1.05)
        categoria = 'Azul'
    
    if categoria == 'Verde':
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Utilizada': list(reversed(fatura_dict['Demanda']['Demanda Verde Medida'])),
            f'Contratada - {dem_c} kW': list(reversed(dem_c_list)),
            f"Proposta - {dem_rec} kW": list(reversed(dem_rec_list)),
            f"Proposta + Tolerância de 5%": list(reversed(lim)),
        }
        estilos_fatura.recomendado_style(dados_dict=dados_dict,fatura_dict=fatura_dict,categoria=categoria,writer=writer,dem_c=dem_c,dem_rec=dem_rec)
    else:
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Utilizada FP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
            'Utilizada P': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
            f'Contratada FP - {dem_c_fp} kW': list(reversed(dem_c_fp_list)),
            f'Contratada P - {dem_c_p} kW': list(reversed(dem_c_p_list)),
            f"Proposta FP - {dem_rec_fp} kW": list(reversed(dem_rec_fp_list)),
            f"Proposta P - {dem_rec_p} kW": list(reversed(dem_rec_p_list)),
            f"Proposta + Tolerância de 5% (FP)": list(reversed(lim_fp)),
            f"Proposta + Tolerância de 5% (P)": list(reversed(lim_p)),
        }
        estilos_fatura.recomendado_style(dados_dict=dados_dict,fatura_dict=fatura_dict,categoria=categoria,writer=writer,dem_c=[dem_c_fp,dem_c_p],dem_rec=[dem_rec_fp,dem_rec_p])