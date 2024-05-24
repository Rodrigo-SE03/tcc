from flask_wtf import FlaskForm
from datetime import datetime
from wtforms import StringField,SubmitField,FloatField,IntegerField,SelectField,FieldList
from wtforms.validators import DataRequired,Length,NumberRange,Regexp

class FormAddCarga(FlaskForm):
    nome_equip = StringField('Nome do equipamento', validators= [DataRequired()])
    potencia = FloatField('Potência (kW)', validators= [DataRequired()])
    fp = FloatField('Fator de potência', validators= [DataRequired()],default=1)
    fp_tipo = SelectField('Tipo de fator de potência', validators= [DataRequired()],choices=['Indutivo','Capacitivo'])
    qtd = IntegerField('Quantidade', validators= [DataRequired()],default=1)
    hr_inicio = StringField('Início', validators= [DataRequired(),Regexp(r"^[0-9][0-9][:][0-9][0-9]",message="O formato deve ser hh:mm")],default="00:00")
    hr_fim = StringField('Fim', validators= [DataRequired(),Regexp(r"^[0-9][0-9][:][0-9][0-9]",message="O formato deve ser hh:mm")],default="00:00")
    dia_sem = SelectField('Dias de Uso', validators= [DataRequired()],choices=['Dias Úteis','Sábados','Domingos'])

    add_button = SubmitField('Adicionar', validators= [DataRequired()])

class FormInfo(FlaskForm):
    dias_u = IntegerField('Nº de dias úteis', validators= [DataRequired(),NumberRange(min=1,max=23,message="Valor inválido")],default=22)
    dias_s = IntegerField('Nº de sábados', validators= [DataRequired(),NumberRange(min=0,max=5,message="Valor inválido")],default=4)
    dias_d = IntegerField('Nº de domingos', validators= [DataRequired(),NumberRange(min=0,max=5,message="Valor inválido")],default=4)
    ponta = FloatField('Início do horário de ponta', validators= [DataRequired(),NumberRange(min=0,max=21,message="Valor inválido")],default=18)

    registrar_info = SubmitField('Registrar', validators= [DataRequired()])

class FormSalvarCargas(FlaskForm):
    nome = StringField('Nome do arquivo', validators= [DataRequired()])
    salvar_btn = SubmitField('Salvar', validators= [DataRequired()])

class SelecionarGrupo(FlaskForm):
    grupo = SelectField('Grupo tarifário desejado', validators= [DataRequired()],choices=['-selecionar-','Grupo B','Grupo A'])
    selecionar = SubmitField('Selecionar', validators= [DataRequired()])

class SelecionarAnalise(FlaskForm):
    tipo = SelectField('Modelo de preenchimento de dados desejado', validators= [DataRequired()],choices=['-selecionar-','Automático','Manual'])
    selecionar = SubmitField('Selecionar', validators= [DataRequired()])

class FormTarifasB(FlaskForm):
    convencional = FloatField('Valor da tarifa convencional (R$/kWh)', validators= [DataRequired()])
    branca_fp = FloatField('Tarifa Branca - horário fora de ponta (R$/kWh)', validators= [DataRequired()])
    branca_i = FloatField('Tarifa Branca - horário intermediário (R$/kWh)', validators= [DataRequired()])
    branca_p = FloatField('Tarifa Branca - horário de ponta (R$/kWh)', validators= [DataRequired()])

    registrar_b = SubmitField('Registrar', validators= [DataRequired()])

class FormTarifasA(FlaskForm):
    verde_fp = FloatField('Valor da tarifa de consumo - horário fora de ponta (R$/kWh)', validators= [DataRequired()])
    verde_p = FloatField('Valor da tarifa de consumo - horário de ponta (R$/kWh)', validators= [DataRequired()])
    verde_dem = FloatField('Valor da tarifa de demanda única (R$/kW)', validators= [DataRequired()])
    
    azul_fp = FloatField('Valor da tarifa de consumo - horário fora de ponta (R$/kWh)', validators= [DataRequired()])
    azul_p = FloatField('Valor da tarifa de consumo - horário de ponta (R$/kWh)', validators= [DataRequired()])
    azul_dem_fp = FloatField('Valor da tarifa de demanda - horário fora de ponta (R$/kW)', validators= [DataRequired()])
    azul_dem_p = FloatField('Valor da tarifa de demanda - horário de ponta (R$/kW)', validators= [DataRequired()])

    te = FloatField('Valor da tarifa de referência reativa - TE do subgrupo B1 (R$/kWh)', validators= [DataRequired()])

    registrar_a = SubmitField('Registrar', validators= [DataRequired()])

class FormFatura(FlaskForm):
    dem_c_fp = IntegerField('Demanda Contratada Fora de Ponta ou Única (kW)', validators= [DataRequired(),NumberRange(min=30,max=15000,message="Valor inválido")])
    dem_c_p = IntegerField('Demanda Contratada na Ponta (kW) - Manter 0 Para Modalidade Verde ',validators=[NumberRange(min=0,max=15000,message="Valor inválido")],default=0)

    reg = SubmitField('Registrar', validators= [DataRequired()])

class FormData(FlaskForm):
    mes = SelectField('Mês da fatura mais recente', validators= [DataRequired()],choices=['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'])
    ano = IntegerField('Ano da fatura mais recente',validators=[NumberRange(min=2000,max=2100,message="Valor inválido"),DataRequired()])
    registrar = SubmitField('Registrar', validators= [DataRequired()])

class FormManual(FlaskForm):
    demanda_p = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    demanda_fp = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    consumo_p = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    consumo_fp = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    consumo_hr = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    ufer = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    ufer_hr = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    dmcr = FieldList(FloatField(validators= [DataRequired(),NumberRange(min=1,message="Valor inválido")]),min_entries=12,max_entries=12)
    registrar_dados = SubmitField('Registrar', validators= [DataRequired()])

class FormSalvarFatura(FlaskForm):
    nome = StringField('Nome do arquivo', validators= [DataRequired()])
    salvar_btn = SubmitField('Salvar', validators= [DataRequired()])
