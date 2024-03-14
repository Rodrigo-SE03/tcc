import pandas as pd
from fatura_dir import graficos

#Definição do estilo visual da tabela de consumo por carga
def comparar_geral_style(worksheet,workbook,verde_dict,azul_dict,fatura_dict,categoria,total):
    merge_format = workbook.add_format({
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#A6A6A6",
        "font_color": "black",
        "border_color": "black"
    })

    mes_format = workbook.add_format({
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "white",
        "font_color": "black",
        "border_color": "black"
    })
    pot_format_v = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B050",
        "font_color": "black",
        "border_color": "black",
        'num_format':'#,##0.00 "kW"'
    })
    energia_format_v = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B050",
        "font_color": "black",
        "border_color": "black",
        'num_format':'#,##0.00 "kWh"'
    })
    custo_format_v = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B050",
        "font_color": "black",
        "border_color": "black",
        'num_format':'R$ #,##0.00'
    })

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
    worksheet.merge_range(1,7,1,8,"FORA DE PONTA",merge_format)
    worksheet.write(2,7,'kWh',merge_format)
    worksheet.write(2,8,'R$',merge_format)
    worksheet.merge_range(0,9,1,9,"TOTAL",merge_format)
    worksheet.write(2,9,'R$',merge_format)

    pot_format_a = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B0F0",
        "font_color": "black",
        "border_color": "black",
        'num_format':'#,##0.00 "kW"'
    })
    energia_format_a = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B0F0",
        "font_color": "black",
        "border_color": "black",
        'num_format':'#,##0.00 "kWh"'
    })
    custo_format_a = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "fg_color": "#00B0F0",
        "font_color": "black",
        "border_color": "black",
        'num_format':'R$ #,##0.00'
    })

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
    worksheet.merge_range(0,12,1,15,"DEMANDA",merge_format)
    worksheet.write(2,12,'kW',merge_format)
    worksheet.write(2,13,'R$',merge_format)
    worksheet.write(2,14,'kW',merge_format)
    worksheet.write(2,15,'R$',merge_format)
    worksheet.merge_range(0,16,1,19,"ULTRAPASSAGEM",merge_format)
    worksheet.write(2,16,'kW',merge_format)
    worksheet.write(2,17,'R$',merge_format)
    worksheet.write(2,18,'kW',merge_format)
    worksheet.write(2,19,'R$',merge_format)
    worksheet.merge_range(0,20,0,23,"CONSUMO",merge_format)
    worksheet.merge_range(1,20,1,21,"PONTA",merge_format)
    worksheet.write(2,20,'kWh',merge_format)
    worksheet.write(2,21,'R$',merge_format)
    worksheet.merge_range(1,22,1,23,"FORA DE PONTA",merge_format)
    worksheet.write(2,22,'kWh',merge_format)
    worksheet.write(2,23,'R$',merge_format)
    worksheet.merge_range(0,24,1,24,"TOTAL",merge_format)
    worksheet.write(2,24,'R$',merge_format)

    worksheet.set_column('B:J',14)
    worksheet.set_column('L:Y',14)

    pot_format = workbook.add_format({"border": 1,"align": "right","valign": "vcenter",'num_format':'#,##0.00 "kW"','border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1,"align": "right","valign": "vcenter"})
    if categoria == 'Verde':
        worksheet.write('M18','Demanda Contratada Atual:')
        worksheet.write('N18',fatura_dict['Demanda']['Demanda Contratada Atual'],pot_format)
        
    else:
        worksheet.write('M18','Demanda Contratada P Atual:')
        worksheet.write('N18',fatura_dict['Demanda']['Demanda Contratada P Atual'],pot_format)
        worksheet.write('M19','Demanda Contratada FP Atual:')
        worksheet.write('N19',fatura_dict['Demanda']['Demanda Contratada FP Atual'],pot_format)

    if total[0]-total[1] <= 0:
        worksheet.write('M21','Demanda Contratada Recomendada:')
        worksheet.write('N22',fatura_dict['Demanda']['Demanda Contratada FP Indicada'],pot_format)
    else:
        worksheet.write('M21','Demanda Contratada P Recomendada:')
        worksheet.write('N21',fatura_dict['Demanda']['Demanda Contratada P Indicada'],pot_format)
        worksheet.write('M22','Demanda Contratada FP Recomendada:')
        worksheet.write('N22',fatura_dict['Demanda']['Demanda Contratada FP Indicada'],pot_format)

    worksheet.write('Q18','Custo - Verde')
    worksheet.write('R18',total[0],rs_format)
    worksheet.write('Q19','Custo - Azul')
    worksheet.write('R19',total[1],rs_format)
    graficos.graf_compara_custos(worksheet=worksheet,workbook=workbook,sheet_name='Comparativo')
#--------------------------------------------------------------------------------------------------------
    
def recomendado_style(dados_dict,fatura_dict,categoria,writer,dem_c,dem_rec):
    dados_df = pd.DataFrame(dados_dict)
    dados_df.to_excel(writer,sheet_name='Recomendação',startrow=1,header=False,index=False)
    workbook = writer.book
    worksheet = writer.sheets["Recomendação"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(0, 0, max_row, max_col-1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 0.1)
    if categoria == 'Verde':

        custo_dict = {
            'Demandas Contratadas': fatura_dict['Demanda']['Lista de demandas contratadas'],
            'Custos Anuais': fatura_dict['Demanda']['Lista de custos anuais por demanda contratada']
        }


        dem_min = min([dem_c,dem_rec])-100
        if dem_min < 0: dem_min = 0
        dem_max = max([dem_rec,dem_c])+100

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
        column_settings = [{"header": column} for column in custo_df.columns]
        worksheet.write(0,last_col+max_col,'Custo')
        worksheet.write(0,last_col+max_col+1,'Demanda Contratada Atual')
        worksheet.write(0,last_col+max_col+2,'Demanda Contratada Recomendada')
        worksheet.write(custo_dict['Demandas Contratadas'].index(dem_rec)+1,last_col+max_col+2,custo_dict['Custos Anuais'][custo_dict['Demandas Contratadas'].index(dem_rec)])
        worksheet.write(custo_dict['Demandas Contratadas'].index(dem_c)+1,last_col+max_col+1,custo_dict['Custos Anuais'][custo_dict['Demandas Contratadas'].index(dem_c)])
        worksheet.add_table(0, last_col+1, max_row, last_col+max_col+2, {"columns": column_settings})

        worksheet.set_column(0, last_col+max_col+2, 0.1)
        graficos.graf_demanda_verde(sheet_name='Recomendação',workbook=workbook,worksheet=worksheet,dem_c=dem_c,dem_rec=dem_rec,custo_dict=custo_dict)