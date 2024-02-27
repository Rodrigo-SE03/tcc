from flask import Flask, render_template, url_for, request, flash, send_from_directory
from forms import FormAddCarga,SelecionarGrupo,FormTarifasB,FormTarifasA,FormSalvarCargas
from werkzeug.utils import secure_filename
import os
import tratar_cargas,tratar_tarifas,planilha

ALLOWED_EXTENSIONS = {'xlsx'}
UPLOAD_FOLDER = 'Planilhas'

h_p = 17
dias = 22
nome_arquivo = ''
tarifas_dict = {
    'convencional': 0.0,
    'branca': [0.0,0.0,0.0],
    'verde':[0.0,0.0,0.0],
    'azul':[0.0,0.0,0.0,0.0]
}
grupo = '-selecionar-'
cargas_dict = {
        'Carga':[],
        'Potência':[],
        'FP':[],
        'FP - Tipo':[],
        'Quantidade':[],
        'Início':[],
        'Fim':[],
        'Remover': []
    }

app = Flask(__name__)

app.config['SECRET_KEY'] = '358823e5046ab23c149ff9a047b30ae8'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/cargas",methods = ['GET','POST'])
def cargas():
    form_add_carga = FormAddCarga()
    form_salvar_cargas = FormSalvarCargas()
    global cargas_dict
    global tarifas_dict
    global grupo
    global nome_arquivo
    global h_p
    global dias
    
    #Procedimento para adicionar a carga ao dicionário principal
    if form_add_carga.validate_on_submit() and 'add_button' in request.form:    
        tratar_cargas.nova_carga(cargas_dict,form_add_carga)  
        print(cargas_dict)
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
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            cargas_dict = tratar_cargas.carregar_cargas(file = filename,folder=UPLOAD_FOLDER)
            flash('Arquivo carregado',category='alert-success')
            return app.redirect(url_for('cargas'))
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para remover uma carga do dicionário principal
    length = len(cargas_dict['Carga'])
    
    if 'remove_btn' in request.form:            
        tratar_cargas.remover_carga(cargas=cargas_dict,i=int(request.form['remove_btn']))   
        return app.redirect(url_for("cargas"))
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para salvar a planilha com as análises
    if 'salvar_btn' in request.form:            
        msg = tratar_cargas.verificar_save(cargas_dict=cargas_dict,tarifas_dict=tarifas_dict)
        if msg != 'Arquivo salvo com sucesso': 
            flash(msg,category='alert-danger')
        else:
            dias = form_add_carga.dias.data
            h_p = form_add_carga.ponta.data
            nome_arquivo = form_salvar_cargas.nome.data
            planilha.limpar_pasta(folder=UPLOAD_FOLDER)
            return app.redirect(url_for("download"))
        return app.redirect(url_for("cargas"))
    #--------------------------------------------------------------------------------------------------------

    return render_template('cargas.html',form_add_carga = form_add_carga,cargas_dict=cargas_dict,length=length,form_salvar_cargas=form_salvar_cargas)


@app.route("/tarifas",methods = ['GET','POST'])
def tarifas():
    global grupo
    global tarifas_dict
    form_selecionar_grupo = SelecionarGrupo()
    
    #Procedimentos para inicialização das variáveis de tarifas
    if grupo == 'Grupo B':          
        form_tarifas_b = FormTarifasB(convencional = tarifas_dict['convencional'],branca_fp=tarifas_dict['branca'][0],branca_i=tarifas_dict['branca'][1],branca_p=tarifas_dict['branca'][2])
        form_tarifas_a = FormTarifasA()
    elif grupo == 'Grupo A':
        form_tarifas_a = FormTarifasA(verde_fp=tarifas_dict['verde'][0],verde_p=tarifas_dict['verde'][1],verde_dem=tarifas_dict['verde'][2],azul_fp=tarifas_dict['azul'][0],azul_p=tarifas_dict['azul'][1],azul_dem_fp=tarifas_dict['azul'][2],azul_dem_p=tarifas_dict['azul'][3])
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
            'azul':[0.0,0.0,0.0,0.0]
        }
        return app.redirect(url_for('tarifas'))
    #--------------------------------------------------------------------------------------------------------
    
    #Procedimento para definir valor das tarifas do Grupo B
    if form_tarifas_b.validate_on_submit() and 'registrar_b' in request.form:   
        tarifas_dict = tratar_tarifas.registrar_tarifas(tarifas=tarifas,form=form_tarifas_b,grupo=grupo)
    #--------------------------------------------------------------------------------------------------------

    #Procedimento para definir valor das tarifas do Grupo A
    if form_tarifas_a.validate_on_submit() and 'registrar_a' in request.form:   
        tarifas_dict = tratar_tarifas.registrar_tarifas(tarifas=tarifas,form=form_tarifas_a,grupo=grupo)
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
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            tarifas_dict = tratar_tarifas.carregar_tarifas(file = filename,folder=UPLOAD_FOLDER,grupo=grupo)
            flash('Tarifas carregadas',category='alert-success')

            return app.redirect(url_for('tarifas'))
    #--------------------------------------------------------------------------------------------------------

    return render_template('tarifas.html',form_selecionar_grupo = form_selecionar_grupo,grupo = grupo,form_tarifas_b=form_tarifas_b,form_tarifas_a=form_tarifas_a)


@app.route('/download')
def download():
    global nome_arquivo
    nome = f'{nome_arquivo}.xlsx'
    planilha.criar_planilha(cargas=cargas_dict,tarifas_dict=tarifas_dict,grupo=grupo,nome=nome,folder=UPLOAD_FOLDER,h_p=h_p,dias=dias)
    uploads = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])
    return send_from_directory(directory=uploads, path=nome)


@app.route("/faturas")
def faturas():
    return render_template('faturas.html')

@app.route("/")
def home():
    return render_template('home.html')

if __name__ == '__main__':
    app.run(debug=True)