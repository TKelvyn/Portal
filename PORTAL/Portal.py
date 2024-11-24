import os
import json
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session, send_from_directory
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
import requests
import base64
from openpyxl import load_workbook
from openpyxl import Workbook


app = Flask(__name__) 

chromedriver_autoinstaller.install()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "Promessa", "Promessa.xlsx")

if not os.path.exists(EXCEL_PATH):
    wb = load_workbook()
    ws = wb.active
    ws.title = "Promessas"
    ws.append(['CNPJ', 'Data Promessa'])
    wb.save(EXCEL_PATH)

SENHAS_ARQUIVO = os.path.join(BASE_DIR, 'templates', 'Senhas.txt')
DATA_JSON_ARQUIVO = os.path.join(BASE_DIR, 'templates', 'data.json')

app.secret_key = os.urandom(24) 

cnpjs = []
data = []
nome = []
venc = []
saldo = []
emails = []

def verificar_senha(token):
    with open(SENHAS_ARQUIVO, 'r') as f:
        senhas = f.readlines()
    
    for linha in senhas:
        if not linha.strip():
            continue
        
        dados = linha.strip().split(',')
        
        if len(dados) != 7:
            continue  
        
        stored_token, cnpj, nome_empresa, data_geracao, vencimento_real, saldo_a_receber, email = dados
        if token == stored_token:
            session['cnpj'] = cnpj
            session['data'] = data_geracao
            session['nome_empresa'] = nome_empresa
            session['vencimento_real'] = vencimento_real
            session['saldo_a_receber'] = saldo_a_receber
            session['email'] = email
            cnpjs.append(cnpj)
            data.append(data_geracao)
            nome.append(nome_empresa)
            venc.append(vencimento_real)
            saldo.append(saldo_a_receber)
            emails.append(email)
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

    if cnpjs:
        def obter_token():
            
            url_token = "http://172.50.30.239:8962/REST/api/oauth2/v1/token"
            payload_token = {
                'grant_type': 'password',
                'username': 'INTEGRACAO',
                'password': 'INTEGRACAO'
            }
            headers_token = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            response_token = requests.post(url_token, data=payload_token, headers=headers_token)

            if response_token.status_code == 201:
                token_data = response_token.json()
                return token_data.get('access_token')
            else:
                print(f"Erro na autenticação: {response_token.status_code}, {response_token.text}")
                return None
            
        def processar_cnpj(cnpj, access_token):
            url_chatbot = "http://172.50.30.239:8962/rest/CHATBOT"
            headers_chatbot = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            params_chatbot = {'CNPJ': cnpj}

            response_chatbot = requests.get(url_chatbot, headers=headers_chatbot, json=params_chatbot)

            if response_chatbot.status_code == 200:
                response_json = response_chatbot.json()
                print(f"Resposta do Chatbot para o CNPJ {cnpj}: {response_json}")
                if isinstance(response_json, list) and len(response_json) > 0:
                    boleto_data = response_json[0]
                    
                    boleto_base64 = boleto_data.get('BOLETOS    ', None)

                    if boleto_base64:
                        try:
                            boleto_bytes = base64.b64decode(boleto_base64)

                            pasta_cnpj = os.path.join(BASE_DIR,"Boletos", str(cnpj))

                            os.makedirs(pasta_cnpj, exist_ok=True)

                            arquivo_destino = os.path.join(pasta_cnpj, f"boleto_{cnpj}.pdf")

                            with open(arquivo_destino, "wb") as f:
                                f.write(boleto_bytes)
                            print(f"Boleto para o CNPJ {cnpj} salvo como {arquivo_destino}")
                        except Exception as e:
                            print(f"Erro ao decodificar o boleto para o CNPJ {cnpj}: {str(e)}")
                    else:
                        print(f"Erro: Boleto não encontrado para o CNPJ {cnpj}.")
            elif response_chatbot.status_code == 500:
                print(f"Erro 500 ao processar o CNPJ {cnpj}. Tentando o próximo.")
            else:
                print(f"Erro ao processar o CNPJ {cnpj}: {response_chatbot.status_code}, {response_chatbot.text}")
        
        def main():
            access_token = obter_token()
            if not access_token:
                return

            cnpj = cnpjs

            print(f"\nIniciando o processamento para o CNPJ {cnpj}...")
            processar_cnpj(cnpj, access_token)
        
        main()

    empresa = nome[0]

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

    df = pd.DataFrame(data)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.expand_frame_repr', False)

    columns_to_check = ['CNPJ']

    for col_name in columns_to_check:
        if col_name not in df.columns:
            df[col_name] = ''

    colunas_pd = ['CNPJ']

    df_filtrado = df[colunas_pd]
    
    data = df_filtrado.to_dict(orient='records')

    if not df_filtrado.empty:
        for registro in data:
            print("Registro:",registro)
            
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
    return jsonify({'empresa': empresa, 'datahora': datahora})

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

@app.route('/gerar_boleto', methods=['POST'])
def gerar_boleto():
    cnpj = request.form.get('cnpj')  
    data_promessa = request.form.get('data_promessa') 

    cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '') 
    pasta_cnpj = os.path.join(BASE_DIR,"Boletos", cnpj)  
    arquivo_boleto = os.path.join(pasta_cnpj, f"boleto_{cnpj}.pdf")  
    print(arquivo_boleto)

    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.append([cnpj, data_promessa])  
    wb.save(EXCEL_PATH)

    if os.path.exists(arquivo_boleto):
        return send_from_directory(pasta_cnpj, f"boleto_{cnpj}.pdf", as_attachment=True)
    else:
        return jsonify({
            'erro': 'Boleto não encontrado. Por favor, entre em contato com o chatbot para obter o boleto.'
        }), 404



if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
