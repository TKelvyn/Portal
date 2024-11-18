import os
import json
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import chromedriver_autoinstaller
import re

app = Flask(__name__) 

chromedriver_autoinstaller.install()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SENHAS_ARQUIVO = os.path.join(BASE_DIR, 'templates', 'Senhas.txt')
DATA_JSON_ARQUIVO = os.path.join(BASE_DIR, 'templates', 'data.json')

app.secret_key = os.urandom(24) 

cnpjs = []

def verificar_senha(token):
    # Carrega os tokens e CNPJs existentes no arquivo
    with open(SENHAS_ARQUIVO, 'r') as f:
        senhas = f.readlines()
    
    for linha in senhas:
        stored_token, cnpj = linha.strip().split(',')
        if token == stored_token:
            session['cnpj'] = cnpj
            cnpjs.append(cnpj)
            senhas.remove(linha)
            with open(SENHAS_ARQUIVO, 'w') as f:
                f.writelines(senhas)
            return True  
    return False 

@app.route('/')
def index():
    return render_template('login.html')

@app.route('/verificar', methods=['POST'])
def verificar():
    if request.method == 'POST':
        token = request.form['senha']  
        if verificar_senha(token):
            return redirect(url_for('loading'))
        else:
            return jsonify({"error": "Token inválido!"}), 403  
    return render_template('login.html')

@app.route('/loading', methods=['GET'])
def loading():
    return render_template('loading.html')

@app.route('/process', methods=['GET'])
def process():
    with open(DATA_JSON_ARQUIVO, 'w') as json_file:
        json_file.write('')

    # Chamada Pro Endpoint do Kimura 

    empresa = "Teste Empresa"

    all_grouped_data = {}

    grouped_data = {}

    data = []

    def format_cnpj(cnpj):
        cnpj = str(cnpj)
        cnpj = re.sub(r'[^0-9]', '', cnpj)  
        return f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
    
    cnpj_formatados = format_cnpj(cnpjs)

    data.append({
        'CNPJ': cnpj_formatados ,
    })

    # Criar um DataFrame
    df = pd.DataFrame(data)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.expand_frame_repr', False)

    columns_to_check = ['CNPJ', 'Data vencimento', 'Valor',]

    for col_name in columns_to_check:
        if col_name not in df.columns:
            df[col_name] = ''

    colunas_pd = ['CNPJ', 'Data vencimento', 'Valor',]

    df_filtrado = df[colunas_pd]
    
    data = df_filtrado.to_dict(orient='records')

    if not df_filtrado.empty:
        for registro in data:
            print("Registro:",registro)
            empresa = empresa 
            if empresa not in grouped_data:
                grouped_data[empresa] = []
            grouped_data[empresa].append(registro)
        
        if grouped_data:
            if empresa not in all_grouped_data:
                all_grouped_data[empresa] = []
            all_grouped_data[empresa].append(grouped_data)

    datahora = datetime.now()  
    datahora = datahora.strftime("%d/%m/%Y %H:%M")

    print(all_grouped_data)
    with open(DATA_JSON_ARQUIVO, 'a') as json_file:
        json.dump(all_grouped_data, json_file)
    return redirect(url_for('empresas', empresa=empresa, datahora=datahora))

@app.route('/data', methods=['GET'])
def get_data():
    with open(DATA_JSON_ARQUIVO, 'r') as json_file:
        data = json.load(json_file)
    
    return jsonify(data)

@app.route('/empresas', methods=['GET'])
def empresas():
    empresa = request.args.get('empresa')
    datahora = request.args.get('datahora')
    print(empresa)
    return render_template('empresas.html', empresa=empresa, datahora=datahora)

@app.route('/gif/<int:gif_number>')
def get_gif(gif_number):
    gif_path = os.path.join(app.root_path, 'templates', 'static', 'images', f'Animation - {gif_number}.GIF')
    return send_file(gif_path, mimetype='image/gif')


if __name__ == '__main__':
    #app.run(debug=True)
    app.run(host='0.0.0.0', port=5000, debug=True)