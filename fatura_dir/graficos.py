def graf_consumo(worksheet,workbook,sheet_name,categoria):
    chart = workbook.add_chart({'type':'column'})
    if categoria == 'Verde':
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "P",'values':f"={sheet_name}!$H$2:$H$13",'fill':{'color':'#C0504D'},'overlap':-20})
        chart.add_series({'name': "FP",'values':f"={sheet_name}!$J$2:$J$13",'fill':{'color':'#4F81BD'}})
    else:
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "P",'values':f"={sheet_name}!$K$2:$K$13",'fill':{'color':'#C0504D'},'overlap':-20})
        chart.add_series({'name': "FP",'values':f"={sheet_name}!$M$2:$M$13",'fill':{'color':'#4F81BD'}})
    chart.set_size({'width': 535.75, 'height': 348.8503937})
    chart.set_y_axis({'num_format': "R$ #,##0.00"})
    chart.set_title({'name':'Custos com Consumo de Energia Ativa'})
    chart.set_legend({'position': 'right'})
    worksheet.insert_chart('H16', chart)

def graf_reativos(worksheet,workbook,sheet_name,categoria):
    chart = workbook.add_chart({'type':'column'})
    if categoria == 'Verde':
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "UFER",'values':f"={sheet_name}!$L$2:$L$13",'fill':{'color':'#31859C'},'overlap':-20})
        chart.add_series({'name': "DMCR",'values':f"={sheet_name}!$N$2:$N$13",'fill':{'color':'#E46C0A'}})
    else:
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "UFER",'values':f"={sheet_name}!$O$2:$O$13",'fill':{'color':'#31859C'},'overlap':-20})
        chart.add_series({'name': "DMCR",'values':f"={sheet_name}!$Q$2:$Q$13",'fill':{'color':'#E46C0A'}})
    chart.set_size({'width': 535.75, 'height': 348.8503937})
    chart.set_y_axis({'num_format': "R$ #,##0.00"})
    chart.set_title({'name':'Custos com Reativos'})
    chart.set_legend({'position': 'right'})
    worksheet.insert_chart('B16', chart)

def graf_ultrapassagem(worksheet,workbook,sheet_name,categoria):
    chart = workbook.add_chart({'type':'column'})
    if categoria == 'Verde':
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "Ultrapassagem",'values':f"={sheet_name}!$F$2:$F$13",'fill':{'color':'#8064A2'},'overlap':-20})
    else:
        chart.add_series({'categories':f"={sheet_name}!$A$2:$A$13",'name': "P",'values':f"={sheet_name}!$G$2:$G$13",'fill':{'color':'#31859C'},'overlap':-20})
        chart.add_series({'name': "FP",'values':f"={sheet_name}!$I$2:$I$13",'fill':{'color':'#8064A2'}})
    chart.set_size({'width': 535.75, 'height': 348.8503937})
    chart.set_y_axis({'num_format': "R$ #,##0.00"})
    chart.set_title({'name':'Multas por Ultrapassagem de Demanda'})
    chart.set_legend({'position': 'right'})
    worksheet.insert_chart('N16', chart)

def graf_compara_custos(worksheet,workbook,sheet_name):
    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='{sheet_name}'!$A$4:$A$15",'name': "Verde",'values':f"='{sheet_name}'!$J$4:$J$15",'line':{'color':'#00B050','width':1.5},'smooth':True})
    chart.add_series({'name': "Azul",'values':f"='{sheet_name}'!$Y$4:$Y$15",'line':{'color':'#0070C0','width':1.5},'smooth':True})

    chart.set_y_axis({'name': 'Custo Mensal Estimado (R$)'}) 
    chart.set_x_axis({'major_gridlines':{'visible':True}})
    chart.set_size({'width': 1071.496063, 'height': 348.8503937}) 
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('B18', chart)


def graf_demanda_verde(workbook,worksheet,sheet_name,dem_c,dem_rec,custo_dict):
    
    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='{sheet_name}'!$A$2:$A$13",'name': f"='{sheet_name}'!$B$1",'values':f"='{sheet_name}'!$B$2:$B$13",'line':{'color':'#0070C0','width':1.5},'smooth':True})
    chart.add_series({'name': f"='{sheet_name}'!$C$1",'values':f"='{sheet_name}'!$C$2:$C$13",'line':{'color':'#C0504D','width':1.5}})
    chart.add_series({'name': f"='{sheet_name}'!$D$1",'values':f"='{sheet_name}'!$D$2:$D$13",'line':{'color':'#00B050','width':1.5}})
    chart.add_series({'name': f"='{sheet_name}'!$E$1",'values':f"='{sheet_name}'!$E$2:$E$13",'line':{'color':'#FF0000','width':1.5,'dash_type': 'long_dash'}})

    chart.set_y_axis({'name': 'Demanda - kW'}) 
    chart.set_title({'name': 'Demanda Contratada X Demanda Recomendada'})
    chart.set_size({'width': 1071.496063, 'height': 348.8503937}) 
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('L3', chart)

    size = len(custo_dict['Custos Anuais'])+1
    chart2 = workbook.add_chart({'type':'line'})
    chart2.add_series({'categories':f"='{sheet_name}'!$G$2:$G${size}",'name': f"='{sheet_name}'!$H$1",'values':f"='{sheet_name}'!$H$2:$H${size}",'line':{'color':'#0070C0','width':1.5},'smooth':True})
    chart2.add_series({'name': f"='{sheet_name}'!$I$1",'values':f"='{sheet_name}'!$I$2:$I${size}",'marker': {'type': 'circle','color':'green','size':8},'line':{'none':True},'data_labels': {'value': True,'position':'above','num_format':'R$ #,##0.00'}})
    chart2.add_series({'name': f"='{sheet_name}'!$J$1",'values':f"='{sheet_name}'!$J$2:$J${size}",'marker': {'type': 'circle','color':'red','size':8},'line':{'none':True},'data_labels': {'value': True,'position':'above','num_format':'R$ #,##0.00'}})

    chart2.set_title({'name': 'Custo Anual X Demanda Contratada'})
    chart2.set_size({'width': 1071.496063, 'height': 348.8503937}) 
    chart2.set_y_axis({'name': 'Custo Anual','num_format': "R$ #,##0.00"})
    chart2.set_x_axis({'name': 'Demanda','num_font':{'rotation':90}})
    chart2.set_legend({'position': 'bottom'})
    chart2.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('L22', chart2)

def graf_demanda_azul(workbook,worksheet,sheet_name):

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='{sheet_name}'!$A$2:$A$13",'name': f"='{sheet_name}'!$B$1",'values':f"='{sheet_name}'!$B$2:$B$13",'line':{'color':'#0070C0','width':1.5},'smooth':True})
    chart.add_series({'name': f"='{sheet_name}'!$D$1",'values':f"='{sheet_name}'!$D$2:$D$13",'line':{'color':'#C0504D','width':1.5}})
    chart.add_series({'name': f"='{sheet_name}'!$F$1",'values':f"='{sheet_name}'!$F$2:$F$13",'line':{'color':'#00B050','width':1.5}})
    chart.add_series({'name': f"='{sheet_name}'!$H$1",'values':f"='{sheet_name}'!$H$2:$H$13",'line':{'color':'#FF0000','width':1.5,'dash_type': 'long_dash'}})

    chart.set_y_axis({'name': 'Demanda - kW'}) 
    chart.set_title({'name': 'Demanda Contratada X Demanda Recomendada (FP)'})
    chart.set_size({'width': 1071.496063, 'height': 348.8503937}) 
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('L3', chart)

    chart2 = workbook.add_chart({'type':'line'})
    chart2.add_series({'categories':f"='{sheet_name}'!$A$2:$A$13",'name': f"='{sheet_name}'!$C$1",'values':f"='{sheet_name}'!$C$2:$C$13",'line':{'color':'#0070C0','width':1.5},'smooth':True})
    chart2.add_series({'name': f"='{sheet_name}'!$E$1",'values':f"='{sheet_name}'!$E$2:$E$13",'line':{'color':'#C0504D','width':1.5}})
    chart2.add_series({'name': f"='{sheet_name}'!$G$1",'values':f"='{sheet_name}'!$G$2:$G$13",'line':{'color':'#00B050','width':1.5}})
    chart2.add_series({'name': f"='{sheet_name}'!$I$1",'values':f"='{sheet_name}'!$I$2:$I$13",'line':{'color':'#FF0000','width':1.5,'dash_type': 'long_dash'}})

    chart2.set_y_axis({'name': 'Demanda - kW'}) 
    chart2.set_title({'name': 'Demanda Contratada X Demanda Recomendada (P)'})
    chart2.set_size({'width': 1071.496063, 'height': 348.8503937}) 
    chart2.set_legend({'position': 'bottom'})
    chart2.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('L23', chart2)