import pandas as pd
from fatura_dir import graficos
import math

merge_style = {
    "bold": 1,
    "border": 1,
    "align": "center",
    "valign": "vcenter",
    "fg_color": "#A6A6A6",
    "font_color": "black",
    "border_color": "black",
    "text_wrap":True
}
total_pot_style = {
    "bold": 1,
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#A6A6A6",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kW"'
}
total_energia_style = {
    "bold": 1,
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#A6A6A6",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kWh"'
}
total_custo_style = {
    "bold": 1,
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#A6A6A6",
    "font_color": "black",
    "border_color": "black",
    'num_format':'R$ #,##0.00'
}
total_ufer_style = {
    "bold": 1,
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#A6A6A6",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kVarh"'
}
mes_style = {
    "bold": 1,
    "border": 1,
    "align": "center",
    "valign": "vcenter",
    "fg_color": "white",
    "font_color": "black",
    "border_color": "black"
}
pot_style_v = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B050",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kW"'
}
energia_style_v = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B050",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kWh"'
}
custo_style_v = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B050",
    "font_color": "black",
    "border_color": "black",
    'num_format':'R$ #,##0.00'
}
ufer_style_v = {"border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B050",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kVarh"'}
pot_style_a = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B0F0",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kW"'
}
energia_style_a = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B0F0",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kWh"'
}
custo_style_a = {
    "border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B0F0",
    "font_color": "black",
    "border_color": "black",
    'num_format':'R$ #,##0.00'
}
ufer_style_a = {"border": 1,
    "align": "right",
    "valign": "vcenter",
    "fg_color": "#00B0F0",
    "font_color": "black",
    "border_color": "black",
    'num_format':'#,##0.00 "kVarh"'}

def geral(worksheet,workbook,dados_dict,categoria,total_row):
    global merge_style
    global mes_style
    global pot_style_v
    global energia_style_v
    global custo_style_v
    global ufer_style_v
    global pot_style_a
    global energia_style_a
    global custo_style_a
    global ufer_style_a
    global total_energia_style
    global total_custo_style
    global total_pot_style
    global total_ufer_style

    header_format = workbook.add_format(merge_style)
    mes_format = workbook.add_format(mes_style)
    total_pot_format = workbook.add_format(total_pot_style)
    total_energia_format = workbook.add_format(total_energia_style)
    total_custo_format = workbook.add_format(total_custo_style)
    total_ufer_format = workbook.add_format(total_ufer_style)
    if categoria == "Verde":
        pot_format = workbook.add_format(pot_style_v)
        energia_format = workbook.add_format(energia_style_v)
        custo_format = workbook.add_format(custo_style_v)
        ufer_format = workbook.add_format(ufer_style_v)
        worksheet.write('P3','Demanda',header_format)
        worksheet.write('Q3',total_row[3],total_custo_format)
        worksheet.write('P4','Ultrapassagem',header_format)
        worksheet.write('Q4',total_row[5],total_custo_format)
        worksheet.write('P5','Consumo',header_format)
        worksheet.write('Q5',total_row[7]+total_row[9],total_custo_format)
        worksheet.write('P6','Reativos',header_format)
        worksheet.write('Q6',total_row[11]+total_row[13],total_custo_format)
        worksheet.set_column('P:Q',14)
    else:
        pot_format = workbook.add_format(pot_style_a)
        energia_format = workbook.add_format(energia_style_a)
        custo_format = workbook.add_format(custo_style_a)
        ufer_format = workbook.add_format(ufer_style_a)
        worksheet.write('S3','Demanda',header_format)
        worksheet.write('T3',total_row[2]+total_row[4],total_custo_format)
        worksheet.write('S4','Ultrapassagem',header_format)
        worksheet.write('T4',total_row[6]+total_row[8],total_custo_format)
        worksheet.write('S5','Consumo',header_format)
        worksheet.write('T5',total_row[10]+total_row[12],total_custo_format)
        worksheet.write('S6','Reativos',header_format)
        worksheet.write('T6',total_row[14]+total_row[16],total_custo_format)
        worksheet.set_column('S:T',14)

    i=0
    for key in dados_dict.keys():
        worksheet.write(0,i,key,header_format)
        if "Custo" in key:
            worksheet.write_column(1,i,dados_dict[key],custo_format)
            worksheet.write(13,i,total_row[i],total_custo_format)
        elif "Demanda" in key or "DMCR" in key or 'Ultrapassagem' in key:
            worksheet.write_column(1,i,dados_dict[key],pot_format)
            worksheet.write(13,i,total_row[i],total_pot_format)
        elif "Consumo" in key:
            worksheet.write_column(1,i,dados_dict[key],energia_format)
            worksheet.write(13,i,total_row[i],total_energia_format)
        elif "UFER" in key:
            worksheet.write_column(1,i,dados_dict[key],ufer_format)
            worksheet.write(13,i,total_row[i],total_ufer_format)
        else:
            worksheet.write_column(1,i,dados_dict[key],mes_format)
            worksheet.write(13,i,total_row[i],header_format)
        i+=1

    worksheet.set_column('B:Q',14)

#Definição do estilo visual da tabela de consumo por carga
def comparar_geral_style(worksheet,workbook,verde_dict,azul_dict,fatura_dict,categoria,total):
    global merge_style
    global mes_style
    global pot_style_v
    global energia_style_v
    global custo_style_v
    global pot_style_a
    global energia_style_a
    global custo_style_a

    merge_format = workbook.add_format(merge_style)
    mes_format = workbook.add_format(mes_style)
    pot_format_v = workbook.add_format(pot_style_v)
    energia_format_v = workbook.add_format(energia_style_v)
    custo_format_v = workbook.add_format(custo_style_v)

    col=0
    for key in verde_dict.keys():
        if key == 'Mês':
            worksheet.write_column(3,col,verde_dict[key],mes_format)
        elif ('Demanda' in key or 'Ultrapassagem' in key) and 'Faturado' not in key:
            worksheet.write_column(3,col,verde_dict[key],pot_format_v)
        elif 'Consumo' in key and 'Faturado' not in key:
            worksheet.write_column(3,col,verde_dict[key],energia_format_v)
        else:
            worksheet.write_column(3,col,verde_dict[key],custo_format_v)
        col+=1
    last_v = col

    worksheet.merge_range(0,0,2,0,"MÊS/ANO",merge_format)
    worksheet.merge_range(0,1,1,2,"DEMANDA",merge_format)
    worksheet.write(2,1,'kW',merge_format)
    worksheet.write(2,2,'R$',merge_format)
    worksheet.merge_range(0,3,1,4,"ULTRAPASSAGEM",merge_format)
    worksheet.write(2,3,'kW',merge_format)
    worksheet.write(2,4,'R$',merge_format)
    worksheet.merge_range(0,5,0,8,"CONSUMO",merge_format)
    worksheet.merge_range(1,5,1,6,"PONTA",merge_format)
    worksheet.write(2,5,'kWh',merge_format)
    worksheet.write(2,6,'R$',merge_format)
    worksheet.merge_range(1,7,1,8,"FORA PONTA",merge_format)
    worksheet.write(2,7,'kWh',merge_format)
    worksheet.write(2,8,'R$',merge_format)
    worksheet.merge_range(0,9,1,9,"TOTAL",merge_format)
    worksheet.write(2,9,'R$',merge_format)

    pot_format_a = workbook.add_format(pot_style_a)
    energia_format_a = workbook.add_format(energia_style_a)
    custo_format_a = workbook.add_format(custo_style_a)

    col+=1
    for key in azul_dict.keys():
        if key == 'Mês':
            worksheet.write_column(3,col,azul_dict[key],mes_format)
        elif ('Demanda' in key or 'Ultrapassagem' in key) and 'Faturado' not in key:
            worksheet.write_column(3,col,azul_dict[key],pot_format_a)
        elif 'Consumo' in key and 'Faturado' not in key:
            worksheet.write_column(3,col,azul_dict[key],energia_format_a)
        else:
            worksheet.write_column(3,col,azul_dict[key],custo_format_a)
        col+=1

    worksheet.merge_range(0,11,2,11,"MÊS/ANO",merge_format)
    worksheet.merge_range(0,12,0,15,"DEMANDA",merge_format)
    worksheet.merge_range(1,12,1,13,"PONTA",merge_format)
    worksheet.write(2,12,'kW',merge_format)
    worksheet.write(2,13,'R$',merge_format)
    worksheet.merge_range(1,14,1,15,"FORA PONTA",merge_format)
    worksheet.write(2,14,'kW',merge_format)
    worksheet.write(2,15,'R$',merge_format)
    worksheet.merge_range(0,16,0,19,"ULTRAPASSAGEM",merge_format)
    worksheet.merge_range(1,16,1,17,"PONTA",merge_format)
    worksheet.write(2,16,'kW',merge_format)
    worksheet.write(2,17,'R$',merge_format)
    worksheet.merge_range(1,18,1,19,"FORA PONTA",merge_format)
    worksheet.write(2,18,'kW',merge_format)
    worksheet.write(2,19,'R$',merge_format)
    worksheet.merge_range(0,20,0,23,"CONSUMO",merge_format)
    worksheet.merge_range(1,20,1,21,"PONTA",merge_format)
    worksheet.write(2,20,'kWh',merge_format)
    worksheet.write(2,21,'R$',merge_format)
    worksheet.merge_range(1,22,1,23,"FORA PONTA",merge_format)
    worksheet.write(2,22,'kWh',merge_format)
    worksheet.write(2,23,'R$',merge_format)
    worksheet.merge_range(0,24,1,24,"TOTAL",merge_format)
    worksheet.write(2,24,'R$',merge_format)

    worksheet.set_column('B:J',14)
    worksheet.set_column('L:Y',14)

    pot_format = workbook.add_format({"border": 1,"align": "right","valign": "vcenter","align": "center",'num_format':'#,##0.00 "kW"','border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1,"align": "right","valign": "vcenter"})
    if categoria == 'Verde':
        worksheet.merge_range(17,12,17,13,'Demanda Contratada Atual',merge_format)
        worksheet.merge_range(18,12,18,13,fatura_dict['Demanda']['Demanda Contratada Atual'],pot_format)
        
    else:
        worksheet.merge_range(17,12,17,13,'Demanda Contratada Atual',merge_format)
        worksheet.write('M19','Ponta',mes_format)
        worksheet.write('N19',fatura_dict['Demanda']['Demanda Contratada P Atual'],pot_format)
        worksheet.write('M20','Fora Ponta',mes_format)
        worksheet.write('N20',fatura_dict['Demanda']['Demanda Contratada FP Atual'],pot_format)

    if total[0]-total[1] <= 0:
        worksheet.merge_range(21,12,21,13,'Demanda Contratada Recomendada',merge_format)
        worksheet.merge_range(22,12,22,13,fatura_dict['Demanda']['Demanda Contratada FP Indicada'],pot_format)
    else:
        worksheet.merge_range(21,12,21,13,'Demanda Contratada Recomendada',merge_format)
        worksheet.write('M23','Ponta',mes_format)
        worksheet.write('N23',fatura_dict['Demanda']['Demanda Contratada P Indicada'],pot_format)
        worksheet.write('M24','Fora Ponta',mes_format)
        worksheet.write('N24',fatura_dict['Demanda']['Demanda Contratada FP Indicada'],pot_format)
    worksheet.set_column('M:N',16.2)

    worksheet.merge_range(17,16,17,17,'Custo Anual',merge_format)
    worksheet.write('Q19','Verde',mes_format)
    worksheet.write('R19',total[0],rs_format)
    worksheet.write('Q20','Azul',mes_format)
    worksheet.write('R20',total[1],rs_format)
    graficos.graf_compara_custos(worksheet=worksheet,workbook=workbook,sheet_name='Comparativo')
#--------------------------------------------------------------------------------------------------------
    
def recomendado_style(dados_dict,fatura_dict,ideal,writer,dem_c,dem_rec,economia):
    global merge_style
    global mes_style
    workbook = writer.book
    merge_format = workbook.add_format(merge_style)
    mes_format = workbook.add_format(mes_style)
    pot_format = workbook.add_format({"border": 1,"align": "right","valign": "vcenter","align": "center",'num_format':'#,##0.00 "kW"','border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1,"align": "center","valign": "vcenter"})
    dados_df = pd.DataFrame(dados_dict)
    dados_df.to_excel(writer,sheet_name='Recomendação',startrow=1,header=False,index=False)
    worksheet = writer.sheets["Recomendação"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(0, 0, max_row, max_col-1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 0.1)
    if ideal == 'Verde':

        custo_dict = {
            'Demandas Contratadas': fatura_dict['Demanda']['Lista de demandas contratadas'],
            'Custos Anuais': fatura_dict['Demanda']['Lista de custos anuais por demanda contratada']
        }

        # print(custo_dict['Demandas Contratadas'])


        dem_min = math.ceil((min([dem_c,dem_rec])-100)/5)*5
        if dem_min < 30: dem_min = 30
        dem_max = math.ceil((max([dem_rec,dem_c])+100)/5)*5

        idx_min = custo_dict['Demandas Contratadas'].index(dem_min)
        idx_max = custo_dict['Demandas Contratadas'].index(dem_max)

        custo_new = custo_dict['Custos Anuais'][idx_min:idx_max]
        dem_new = custo_dict['Demandas Contratadas'][idx_min:idx_max]

        custo_dict = {
            'Demandas Contratadas': dem_new,
            'Custos Anuais': custo_new
        }
        
        custo_df = pd.DataFrame(custo_dict)
        last_col = max_col
        custo_df.to_excel(writer,sheet_name='Recomendação',startrow=1,startcol=max_col+1,header=False,index=False)
        (max_row, max_col) = custo_df.shape
        column_settings = [{"header": column} for column in custo_df.columns]                                               #POSSO CORRIGIR ISSO
        worksheet.write(0,last_col+max_col,'Custo')
        worksheet.write(0,last_col+max_col+1,'Demanda Contratada Atual')
        worksheet.write(0,last_col+max_col+2,'Demanda Contratada Recomendada')
        worksheet.write(custo_dict['Demandas Contratadas'].index(dem_rec)+1,last_col+max_col+2,custo_dict['Custos Anuais'][custo_dict['Demandas Contratadas'].index(dem_rec)])
        worksheet.write(custo_dict['Demandas Contratadas'].index(math.ceil(dem_c/5)*5)+1,last_col+max_col+1,custo_dict['Custos Anuais'][custo_dict['Demandas Contratadas'].index(math.ceil(dem_c/5)*5)])
        worksheet.add_table(0, last_col+1, max_row, last_col+max_col+2, {"columns": column_settings})

        worksheet.set_column(0, last_col+max_col+2, 0.1)
        graficos.graf_demanda_verde(sheet_name='Recomendação',workbook=workbook,worksheet=worksheet,dem_c=dem_c,dem_rec=dem_rec,custo_dict=custo_dict)

        worksheet.merge_range(20,12,20,18,"Demanda Contratada Recomendada",merge_format)
        worksheet.merge_range(21,12,21,18,dem_rec,pot_format)

        worksheet.merge_range(20,20,20,26,"Economia Estimada",merge_format)
        worksheet.merge_range(21,20,21,26,economia,rs_format)     
    else:
        worksheet.merge_range(20,12,20,18,"Demanda Contratada Recomendada",merge_format)
        worksheet.merge_range(21,12,21,15,"Ponta",pot_format)
        worksheet.merge_range(21,16,21,18,fatura_dict['Demanda']['Demanda Contratada P Indicada'],pot_format)
        worksheet.merge_range(22,12,22,15,"Fora Ponta",pot_format)
        worksheet.merge_range(22,16,22,18,fatura_dict['Demanda']['Demanda Contratada FP Indicada'],pot_format)

        worksheet.merge_range(20,20,20,26,"Economia Estimada",merge_format)
        worksheet.merge_range(21,20,22,26,economia,rs_format)    

        graficos.graf_demanda_azul(sheet_name='Recomendação',workbook=workbook,worksheet=worksheet)