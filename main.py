from flask import Flask, render_template, url_for, request, flash, send_from_directory
from forms import FormAddCarga,SelecionarGrupo,FormTarifasB,FormTarifasA,FormSalvarCargas,FormFatura,FormSalvarFatura,FormInfo,SelecionarAnalise,FormManual,FormData
from werkzeug.utils import secure_filename
import os
from cargas_dir import planilha_cargas,tratar_cargas
from tarifas_dir import tratar_tarifas
from fatura_dir import planilha_fatura,tratar_fatura
from datetime import datetime


UPLOAD_FOLDER = 'arquivos'

download_flag = ''
h_p = 18.0
dias = {
    'dias_u': 22,
    'dias_s': 4,
    'dias_d': 4
}
data = ['JAN',datetime.today().year]
nome_arquivo = ''
tarifas_dict = {
    'convencional': 0.0,
    'branca': [0.0,0.0,0.0],
    'verde':[0.0,0.0,0.0],
    'azul':[0.0,0.0,0.0,0.0],
    'te': 0.0
}
grupo = '-selecionar-'
tipo_analise = '-selecionar-'
cargas_dict = {
        'Carga':[],
        'Potência (kW)':[],
        'FP':[],
        'FP - Tipo':[],
        'Quantidade':[],
        'Início':[],
        'Fim':[],
        'Dias de Uso': [],
        'Remover': []
    }

fatura_dict = {}
historico_dict={
        'mes':[],
        'ano':[],
        'demanda_p':[],
        'demanda_fp':[],
        'dmcr':[],
        'consumo_p': [],
        'consumo_fp':[],
        'consumo_hr':[],
        'ufer': [],
        'ufer_hr': []
    }
dem_c = 0
meses = ['DEZ','NOV','OUT','SET','AGO','JUL','JUN','MAI','ABR','MAR','FEV','JAN']
anos = [datetime.today().year]*12
app = Flask(__name__)

app.config['SECRET_KEY'] = '358823e5046ab23c149ff9a047b30ae8'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


#Exclusão de arquivos
def limpar_pasta(folder):
    for file in os.listdir(folder):
        if 'modelo' not in file and '.png' not in file:
            os.remove(f'{folder}/{file}')
#--------------------------------------------------------------------------------------------------------

def allowed_file(filename,extension):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in extension


#Página da Análise por Cargas
@app.route("/cargas",methods = ['GET','POST'])
def cargas():
    global cargas_dict
    global tarifas_dict
    global grupo
    global nome_arquivo
    global h_p
    global dias
    global download_flag
    form_add_carga = FormAddCarga()
    form_salvar_cargas = FormSalvarCargas()
    form_info = FormInfo(data = {'ponta':h_p,'dias_u':dias['dias_u'],'dias_s':dias['dias_s'],'dias_d':dias['dias_d']})
    
    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
    
    #Procedimento para adicionar a carga ao dicionário principal
    if form_add_carga.validate_on_submit() and 'add_button' in request.form:    
        tratar_cargas.nova_carga(cargas_dict,form_add_carga)  
        flash(f'Carga adicionada - {form_add_carga.nome_equip.data}',category='alert-success')  #Mensagem de alerta
        return app.redirect(url_for("cargas"))  #Precisa de dar redirect pra não ativar o submit ao recarregar a página
    #--------------------------------------------------------------------------------------------------------
    
    #Procedimento para carregar arquivo com lista pré definida de cargas
    if request.method == 'POST' and 'load_btn' in request.form:     
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename,['xlsx']):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.root_path,app.config['UPLOAD_FOLDER'], filename))
            cargas_dict = tratar_cargas.carregar_cargas(file = filename,folder=os.path.join(app.root_path,UPLOAD_FOLDER))
            if cargas_dict == 'Arquivo inválido':
                flash(cargas_dict,category='alert-danger')
                cargas_dict = {
                    'Carga':[],
                    'Potência (kW)':[],
                    'FP':[],
                    'FP - Tipo':[],
                    'Quantidade':[],
                    'Início':[],
                    'Fim':[],
                    'Remover': []
                }
            else:
                flash('Arquivo carregado',category='alert-success')
            return app.redirect(url_for('cargas'))
        else:
            flash('Formato de arquivo inválio. Deve ser um arquivo .xlsx',category='alert-danger')
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para remover uma carga do dicionário principal
    length = len(cargas_dict['Carga'])
    
    if 'remove_btn' in request.form:          
        removida = tratar_cargas.remover_carga(cargas=cargas_dict,i=int(request.form['remove_btn'])) 
        flash(f'Carga removida - {removida}',category='alert-warning')    
        return app.redirect(url_for("cargas"))
    #--------------------------------------------------------------------------------------------------------

    if 'registrar_info' in request.form:
        h_p = form_info.ponta.data
        dias = {
            'dias_u': form_info.dias_u.data,
            'dias_s': form_info.dias_s.data,
            'dias_d': form_info.dias_d.data
        }
        
        flash('Informações registradas com sucesso',category='alert-success')
        return app.redirect(url_for("cargas"))

    #Procedimento para salvar a planilha com as análises
    if 'salvar_btn' in request.form:            
        limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
        msg = tratar_cargas.verificar_save(cargas_dict=cargas_dict,tarifas_dict=tarifas_dict,h_p=h_p,dias=dias)
        if msg != 'Arquivo salvo com sucesso': 
            flash(msg,category='alert-danger')
        else:
            nome_arquivo = form_salvar_cargas.nome.data
            download_flag = 'Cargas'
            return app.redirect(url_for("download"))
        return app.redirect(url_for("cargas"))
    #--------------------------------------------------------------------------------------------------------

    return render_template('cargas.html',form_add_carga = form_add_carga,cargas_dict=cargas_dict,length=length,form_salvar_cargas=form_salvar_cargas,form_info=form_info)
#--------------------------------------------------------------------------------------------------------


#Página de Tarifas Praticadas
@app.route("/tarifas",methods = ['GET','POST'])
def tarifas():
    global grupo
    global tarifas_dict
    form_selecionar_grupo = SelecionarGrupo(data={'grupo':grupo})
    
    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))

    #Procedimentos para inicialização das variáveis de tarifas
    if grupo == 'Grupo B':          
        form_tarifas_b = FormTarifasB(convencional = tarifas_dict['convencional'],branca_fp=tarifas_dict['branca'][0],branca_i=tarifas_dict['branca'][1],branca_p=tarifas_dict['branca'][2])
        form_tarifas_a = FormTarifasA()
    elif grupo == 'Grupo A':
        form_tarifas_a = FormTarifasA(verde_fp=tarifas_dict['verde'][0],verde_p=tarifas_dict['verde'][1],verde_dem=tarifas_dict['verde'][2],azul_fp=tarifas_dict['azul'][0],azul_p=tarifas_dict['azul'][1],azul_dem_fp=tarifas_dict['azul'][2],azul_dem_p=tarifas_dict['azul'][3],te=tarifas_dict['te'])
        form_tarifas_b = FormTarifasB()
    else:
        form_tarifas_b = FormTarifasB()
        form_tarifas_a = FormTarifasA()
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para reiniciar valor das tarifas caso haja mudança
    if request.method == 'POST' and 'selecionar' in request.form:   
        grupo = form_selecionar_grupo.grupo.data
        tarifas_dict = {
            'convencional': 0.0,
            'branca': [0.0,0.0,0.0],
            'verde':[0.0,0.0,0.0],
            'azul':[0.0,0.0,0.0,0.0],
            'te': 0.0
        }
        return app.redirect(url_for('tarifas'))
    #--------------------------------------------------------------------------------------------------------
    
    #Procedimento para definir valor das tarifas do Grupo B
    if form_tarifas_b.validate_on_submit() and 'registrar_b' in request.form:   
        tarifas_dict = tratar_tarifas.registrar_tarifas(tarifas=tarifas,form=form_tarifas_b,grupo=grupo)
        flash('Tarifas registradas',category='alert-success')
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para definir valor das tarifas do Grupo A
    if form_tarifas_a.validate_on_submit() and 'registrar_a' in request.form:   
        tarifas_dict = tratar_tarifas.registrar_tarifas(tarifas=tarifas,form=form_tarifas_a,grupo=grupo)
        flash('Tarifas registradas',category='alert-success')
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para carregar arquivo com valores pré definidos de tarifas
    if request.method == 'POST' and ('load_tb' in request.form or  'load_ta' in request.form):      
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename,'xlsx'):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.root_path,app.config['UPLOAD_FOLDER'], filename))
            tarifas_dict = tratar_tarifas.carregar_tarifas(file = filename,folder=os.path.join(app.root_path,UPLOAD_FOLDER),grupo=grupo)
            if tarifas_dict == 'Arquivo inválido':
                flash(tarifas_dict,category='alert-danger')
                tarifas_dict = {
                    'convencional': 0.0,
                    'branca': [0.0,0.0,0.0],
                    'verde':[0.0,0.0,0.0],
                    'azul':[0.0,0.0,0.0,0.0],
                    'te': 0.0
                }
            else:
                flash('Tarifas carregadas',category='alert-success')

            return app.redirect(url_for('tarifas'))
        else:
            flash('Formato de arquivo inválio. Deve ser um arquivo .xlsx',category='alert-danger')
    #--------------------------------------------------------------------------------------------------------

    return render_template('tarifas.html',form_selecionar_grupo = form_selecionar_grupo,grupo = grupo,form_tarifas_b=form_tarifas_b,form_tarifas_a=form_tarifas_a)
#--------------------------------------------------------------------------------------------------------


#Página de Análise por Fatura
@app.route("/faturas",methods = ['GET','POST'])
def faturas():
    global grupo
    global download_flag
    global tarifas_dict
    global fatura_dict
    global dem_c
    global nome_arquivo
    global tipo_analise
    global data
    global meses
    global anos
    global historico_dict
    if len(historico_dict['mes']) == 0:
        form_manual = FormManual()
    else:
        form_manual = FormManual(demanda_p = historico_dict['demanda_p'],demanda_fp = historico_dict['demanda_fp'],consumo_p = historico_dict['consumo_p'],consumo_fp = historico_dict['consumo_fp'],consumo_hr = historico_dict['consumo_hr'],
                                 ufer = historico_dict['ufer'],ufer_hr = historico_dict['ufer_hr'],dmcr = historico_dict['dmcr'])
    form_salvar_fatura = FormSalvarFatura()
    form_data = FormData(mes = data[0],ano=data[1])
    form_selecionar_analise = SelecionarAnalise(data={'tipo_analise':tipo_analise})

    #Procedimento para definir o tipo de análise (manual ou automática)
    if request.method == 'POST' and 'selecionar' in request.form:   
        tipo_analise = form_selecionar_analise.tipo.data
        return app.redirect(url_for('faturas'))
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para definir a data da fatura mais recente
    if  request.method == 'POST' and 'registrar' in request.form:   
        data = [form_data.mes.data,form_data.ano.data]
        meses,anos = tratar_fatura.definir_meses(data=data)
        flash(f'Fatura mais recente - {data[0]}/{data[1]}',category='alert-success')
        
        return app.redirect(url_for('faturas'))
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para registrar os dados das faturas anteriores
    if  request.method == 'POST' and 'registrar_dados' in request.form:   
        fatura_dict,historico_dict,data = tratar_fatura.dados_manual(dem_c=dem_c,form_manual=form_manual,tarifas=tarifas_dict,meses=meses,anos=anos)
        meses,anos = tratar_fatura.definir_meses(data=data)
        flash('Dados Registrados',category='alert-success')
        return app.redirect(url_for('faturas'))
    #--------------------------------------------------------------------------------------------------------

    if isinstance(dem_c,int) or isinstance(dem_c,float):
        form_fatura = FormFatura(data = {'dem_c_fp':dem_c,'dem_c_p':0})
    else:
        form_fatura = FormFatura(data = {'dem_c_fp':dem_c[0],'dem_c_p':dem_c[1]})
    
    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
    
    if form_fatura.validate_on_submit() and 'reg' in request.form:   
        dem_c = tratar_fatura.demanda_contratada(form_fatura=form_fatura)
        tipo_analise = '-selecionar-'
        flash('Demandas registradas',category='alert-success')

    #Procedimento para carregar arquivo com valores pré definidos de tarifas
    if request.method == 'POST' and 'load_btn' in request.form:      
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename,'pdf'):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.root_path,app.config['UPLOAD_FOLDER'], filename))
            fatura_dict = tratar_fatura.ler_fatura(file = filename,folder=os.path.join(app.root_path,UPLOAD_FOLDER),tarifas=tarifas_dict,dem_c=dem_c)
            if fatura_dict == 'Arquivo inválido':
                flash(fatura_dict,category='alert-danger')
                fatura_dict = {}
            else:
                flash('Fatura carregada',category='alert-success')
            return app.redirect(url_for('faturas'))
        else:
            flash('Formato de arquivo inválio. Deve ser um arquivo .pdf',category='alert-danger')
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para carregar arquivo excel com o histórico das faturas
    if request.method == 'POST' and 'load' in request.form:      
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename,'xlsx'):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.root_path,app.config['UPLOAD_FOLDER'], filename))
            fatura_dict,historico_dict,data = tratar_fatura.ler_excel(file = filename,folder=os.path.join(app.root_path,UPLOAD_FOLDER),tarifas=tarifas_dict,dem_c=dem_c)
            if fatura_dict == 'Arquivo inválido':
                flash(fatura_dict,category='alert-danger')
                fatura_dict = {}
            else:
                meses,anos = tratar_fatura.definir_meses(data=data)
                flash('Fatura carregada',category='alert-success')
            return app.redirect(url_for('faturas'))
        else:
            flash('Formato de arquivo inválio. Deve ser um arquivo .pdf',category='alert-danger')
    #--------------------------------------------------------------------------------------------------------
        
    #Procedimento para salvar a planilha com as análises
    if 'salvar_btn' in request.form: 
        limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))           
        msg = tratar_fatura.verificar_save(fatura_dict=fatura_dict)
        if msg != 'Arquivo salvo com sucesso': 
            flash(msg,category='alert-danger')
        else:
            nome_arquivo = form_salvar_fatura.nome.data
            download_flag = 'Fatura'
            return app.redirect(url_for("download"))
        return app.redirect(url_for('faturas'))
    #--------------------------------------------------------------------------------------------------------
    return render_template('faturas.html',tarifas_dict = tarifas_dict,form_fatura=form_fatura,form_salvar_fatura=form_salvar_fatura,dem_c=dem_c,fatura_dict=fatura_dict,
                           grupo=grupo,form_selecionar_analise=form_selecionar_analise,tipo_analise=tipo_analise,form_manual=form_manual,form_data=form_data,meses=meses,anos=anos)
#--------------------------------------------------------------------------------------------------------

#Função para download dos resultados
@app.route('/download')
def download():
    global nome_arquivo
    global download_flag
    global fatura_dict
    nome = f'{nome_arquivo}.xlsx'
    if download_flag == 'Cargas':
        planilha_cargas.criar_planilha(cargas=cargas_dict,tarifas_dict=tarifas_dict,grupo=grupo,nome=nome,folder=os.path.join(app.root_path,UPLOAD_FOLDER),h_p=h_p,dias=dias)
    elif download_flag == 'Fatura':
        planilha_fatura.criar_planilha(fatura_dict=fatura_dict,nome=nome,folder=os.path.join(app.root_path,UPLOAD_FOLDER))
    else:
        nome = nome = f'{nome_arquivo}.zip'
    uploads = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])
    return send_from_directory(directory=uploads, path=nome)
#--------------------------------------------------------------------------------------------------------

#Página inicial
@app.route("/")
def home():
    global dem_c
    return render_template('home.html')
#--------------------------------------------------------------------------------------------------------

#Função para resetar os valores
@app.route("/reset")
def reset():
    global download_flag
    global dem_c
    global fatura_dict
    global tarifas_dict
    global cargas_dict
    global h_p
    global dias
    global nome_arquivo
    global grupo 
    global tipo_analise
    global data
    global historico_dict
    global meses
    global anos
    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))


    historico_dict={
        'mes':[],
        'ano':[],
        'demanda_p':[],
        'demanda_fp':[],
        'dmcr':[],
        'consumo_p': [],
        'consumo_fp':[],
        'consumo_hr':[],
        'ufer': [],
        'ufer_hr': []
    }
    download_flag = ''
    h_p = 18.0
    data = ['JAN',datetime.today().year]
    dias = {
        'dias_u': 22,
        'dias_s': 4,
        'dias_d': 4
    }
    meses = ['DEZ','NOV','OUT','SET','AGO','JUL','JUN','MAI','ABR','MAR','FEV','JAN']
    anos = [datetime.today().year]*12
    nome_arquivo = ''
    tarifas_dict = {
        'convencional': 0.0,
        'branca': [0.0,0.0,0.0],
        'verde':[0.0,0.0,0.0],
        'azul':[0.0,0.0,0.0,0.0],
        'te': 0.0
    }
    grupo = '-selecionar-'
    tipo_analise = '-selecionar-'
    cargas_dict = {
            'Carga':[],
            'Potência (kW)':[],
            'FP':[],
            'FP - Tipo':[],
            'Quantidade':[],
            'Início':[],
            'Fim':[],
            'Dias de Uso': [],
            'Remover': []
        }

    fatura_dict = {}
    dem_c = 0
    flash('Valores resetados',category='alert-success')
    return app.redirect(url_for('home'))
#--------------------------------------------------------------------------------------------------------

#Função para baixar os modelos
@app.route("/modelos")
def modelos():
    global nome_arquivo
    global download_flag
    download_flag = 'Modelos'
    nome_arquivo = 'modelos'
    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
    return app.redirect(url_for("download"))
#--------------------------------------------------------------------------------------------------------

if __name__ == '__main__':
    app.run(debug=True)
