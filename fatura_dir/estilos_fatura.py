import pandas as pd

merge_style ={
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
        "border_color": "white"
    }

merge_style2 = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
    }
    
empty_style = {
        "bold": 1,
        "right": 1,
        "fg_color": "#4f81bd",
        "border_color": "white"
    }

#Definição do estilo visual da tabela de consumo por carga
def consumo_equip_style(worksheet,workbook,grupo):
    global merge_style
    global empty_style
    merge_format = workbook.add_format(merge_style)
    empty_format = workbook.add_format(empty_style)
    if grupo == "Grupo A":
        worksheet.merge_range("C1:E1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("F1:H1","Consumo Diário Típico (kWh)",merge_format)
    else:
        worksheet.merge_range("C1:F1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("G1:J1","Consumo Diário Típico (kWh)",merge_format)
    worksheet.write_blank(0,0,'',empty_format)
    worksheet.write_blank(0,1,'',empty_format)
#--------------------------------------------------------------------------------------------------------