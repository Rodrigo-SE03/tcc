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
    total_modalidade = tab_atual(fatura_dict=fatura_dict,writer=writer,categoria=categoria)
    results = tab_analise(fatura_dict=fatura_dict,writer=writer,categoria=categoria,total_modalidade=total_modalidade)
    ideal = results[0]
    economia = results[1]
    tab_resultados(fatura_dict=fatura_dict,writer=writer,ideal=ideal,economia=economia)
    writer.close()
#--------------------------------------------------------------------------------------------------------
    
#Criação da aba com as informações atuais de demanda e consumo da unidade consumidora
def tab_atual(fatura_dict,writer,categoria):

    if categoria == 'Verde':
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Demanda Registrada na Ponta': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
            'Demanda Registrada Fora Ponta': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
            'Custo da Demanda': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda Verde (atual)'])),
            'Demanda Não Utilizada': list(reversed(fatura_dict['Demanda']['Demanda Verde Não Utilizada'])),
            'Custo com Demanda Não Utilizada': list(reversed(fatura_dict['Demanda']['Custos com Demanda Verde Não Utilizada'])),
            'Ultrapassagem Registrada': list(reversed(fatura_dict['Demanda']['Demanda Verde Ultrapassada (atual)'])),
            'Custo da Ultrapassagem': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda Verde (atual)'])),
            'Consumo na Ponta': list(reversed(fatura_dict['Consumo']['Consumo P'])),
            'Custo do Consumo na Ponta': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Verde'])),
            'Consumo Fora Ponta': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
            'Custo do Consumo Fora Ponta': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Verde'])),
            'UFER': list(reversed(fatura_dict['Reativo']['Consumo Reativo Medido'])),
            'Custo da UFER': list(reversed(fatura_dict['Reativo']['Custo do Consumo Reativo'])),
            'DMCR': list(reversed(fatura_dict['Reativo']['Demanda Reativa Medida'])),
            'Custo da DMCR': list(reversed(fatura_dict['Reativo']['Custo da Demanda Reativa'])),
        }
        total_row = ['Total','-','-',sum(dados_dict['Custo da Demanda']),'-',sum(dados_dict['Custo com Demanda Não Utilizada']),
                     '-',sum(dados_dict['Custo da Ultrapassagem']),sum(dados_dict['Consumo na Ponta']),sum(dados_dict['Custo do Consumo na Ponta']),
                     sum(dados_dict['Consumo Fora Ponta']),sum(dados_dict['Custo do Consumo Fora Ponta']),sum(dados_dict['UFER']),sum(dados_dict['Custo da UFER']),
                     '-',sum(dados_dict['Custo da DMCR'])]
        total_modalidade = total_row[3]+total_row[5]+total_row[7]+total_row[9]+total_row[11]
    else:
        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Demanda Registrada na Ponta': list(reversed(fatura_dict['Demanda']['Demanda P Medida'])),
            'Custo da Demanda na Ponta': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda P (atual)'])),
            'Demanda Não Utilizada na Ponta': list(reversed(fatura_dict['Demanda']['Demanda P Não Utilizada'])),
            'Custo com Demanda Não Utilizada na Ponta': list(reversed(fatura_dict['Demanda']['Custos com Demanda P Não Utilizada'])),
            'Demanda Registrada Fora Ponta': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
            'Custo da Demanda Fora Ponta': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda FP (atual)'])),
            'Demanda Não Utilizada Fora Ponta': list(reversed(fatura_dict['Demanda']['Demanda FP Não Utilizada'])),
            'Custo com Demanda Não Utilizada Fora Ponta': list(reversed(fatura_dict['Demanda']['Custos com Demanda FP Não Utilizada'])),
            'Ultrapassagem Registrada na Ponta': list(reversed(fatura_dict['Demanda']['Demanda P Ultrapassada (atual)'])),
            'Custo da Ultrapassagem na Ponta': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P (atual)'])),
            'Ultrapassagem Registrada Fora Ponta': list(reversed(fatura_dict['Demanda']['Demanda FP Ultrapassada (atual)'])),
            'Custo da Ultrapassagem Fora Ponta': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda FP (atual)'])),
            'Consumo na Ponta': list(reversed(fatura_dict['Consumo']['Consumo P'])),
            'Custo do Consumo na Ponta': list(reversed(fatura_dict['Consumo']['Custo Consumo - P Azul'])),
            'Consumo Fora Ponta': list(reversed(fatura_dict['Consumo']['Consumo FP'])),
            'Custo do Consumo Fora Ponta': list(reversed(fatura_dict['Consumo']['Custo Consumo - FP Azul'])),
            'UFER': list(reversed(fatura_dict['Reativo']['Consumo Reativo Medido'])),
            'Custo da UFER': list(reversed(fatura_dict['Reativo']['Custo do Consumo Reativo'])),
            'DMCR': list(reversed(fatura_dict['Reativo']['Demanda Reativa Medida'])),
            'Custo da DMCR': list(reversed(fatura_dict['Reativo']['Custo da Demanda Reativa'])),
        }
        total_row = ['Total','-',sum(dados_dict['Custo da Demanda na Ponta']),'-',sum(dados_dict['Custo com Demanda Não Utilizada na Ponta']),'-',sum(dados_dict['Custo da Demanda Fora Ponta']),'-',sum(dados_dict['Custo com Demanda Não Utilizada Fora Ponta']),
                     '-',sum(dados_dict['Custo da Ultrapassagem na Ponta']),'-',sum(dados_dict['Custo da Ultrapassagem Fora Ponta']),
                     sum(dados_dict['Consumo na Ponta']),sum(dados_dict['Custo do Consumo na Ponta']),sum(dados_dict['Consumo Fora Ponta']),sum(dados_dict['Custo do Consumo Fora Ponta']),
                     sum(dados_dict['UFER']),sum(dados_dict['Custo da UFER']),'-',sum(dados_dict['Custo da DMCR'])]
        total_modalidade = total_row[2]+total_row[4]+total_row[6]+total_row[8]+total_row[10]+total_row[12]+total_row[14]+total_row[16]

    workbook = writer.book
    sheet_name = 'Dados_Atuais'
    worksheet = workbook.add_worksheet(sheet_name)

    estilos_fatura.geral(workbook=workbook,worksheet=worksheet,categoria=categoria,dados_dict=dados_dict,total_row=total_row)
    graficos.graf_reativos(categoria=categoria,sheet_name=sheet_name,workbook=workbook,worksheet=worksheet)
    graficos.graf_consumo(categoria=categoria,sheet_name=sheet_name,workbook=workbook,worksheet=worksheet)
    graficos.graf_ultrapassagem(categoria=categoria,sheet_name=sheet_name,workbook=workbook,worksheet=worksheet)
    graficos.graf_composicao(categoria=categoria,sheet_name=sheet_name,workbook=workbook,worksheet=worksheet)
    return total_modalidade
#--------------------------------------------------------------------------------------------------------

#Criação da aba com a comparação entre as modalidades azul e verde para a unidade consumidora, considerando as demandas contratadas ideais calculadas
def tab_analise(fatura_dict,writer,categoria,total_modalidade):
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
        'Demanda HFP': list(reversed(fatura_dict['Demanda']['Demanda FP Medida'])),
        'Demanda HFP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Demanda - Demanda FP'])),
        'Ultrapassagem HP': list(reversed(fatura_dict['Demanda']['Demanda P Ultrapassada'])),
        'Ultrapassagem HP - Faturado': list(reversed(fatura_dict['Demanda']['Custos com Ultrapassagem - Demanda P'])),
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
    ideal = 'Verde' if total[0]<total[1] else 'Azul'
    economia = total_modalidade - (total[0] if ideal == 'Verde' else total[1])
    results = [ideal,economia]
    return results
#--------------------------------------------------------------------------------------------------------

#Criação da aba com informações da estimativa de gastos e economia da unidade de acordo com a modalidade e demanda contratada indicada
def tab_resultados(fatura_dict,writer,ideal,economia):
    categoria = 0
    print(ideal)
    if 'Demanda Verde Ultrapassada (atual)' in fatura_dict['Demanda'].keys():
        categoria = 'Verde'
    else:                                                                             
        categoria = 'Azul'
    
    if ideal == 'Verde':
        if categoria == 'Azul':
            dem_c = fatura_dict['Demanda']['Demanda Contratada FP Atual']
        else:
            dem_c = fatura_dict['Demanda']['Demanda Contratada Atual']
        dem_rec = fatura_dict['Demanda']['Demanda Contratada FP Indicada']
        dem_rec_list = []
        dem_c_list = []
        lim = []
        for m in fatura_dict['Mês']:
            dem_c_list.append(dem_c)
            dem_rec_list.append(dem_rec)
            lim.append(dem_rec*1.05)

        dados_dict = {
            'Mês': list(reversed(fatura_dict['Mês'])),
            'Utilizada': list(reversed(fatura_dict['Demanda']['Demanda Verde Medida'])),
            f'Contratada - {dem_c} kW': list(reversed(dem_c_list)),
            f"Proposta - {dem_rec} kW": list(reversed(dem_rec_list)),
            f"Proposta + Tolerância de 5%": list(reversed(lim)),
        }
        estilos_fatura.resultados_style(dados_dict=dados_dict,fatura_dict=fatura_dict,ideal=ideal,writer=writer,dem_c=dem_c,dem_rec=dem_rec,economia=economia)

    else:
        if categoria == 'Verde':
            dem_c_fp = fatura_dict['Demanda']['Demanda Contratada Atual']
            dem_c_p = fatura_dict['Demanda']['Demanda Contratada Atual']
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
        estilos_fatura.resultados_style(dados_dict=dados_dict,fatura_dict=fatura_dict,ideal=ideal,writer=writer,dem_c=[dem_c_fp,dem_c_p],dem_rec=[dem_rec_fp,dem_rec_p],economia=economia)