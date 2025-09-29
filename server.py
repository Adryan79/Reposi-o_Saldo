from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
import subprocess # AJUSTE: Importar a biblioteca subprocess
import sys # AJUSTE: Importar a biblioteca sys

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def sanitize_dataframe(df):
    for col in df.select_dtypes(include=['datetime64[ns]']).columns:
        df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')
    df = df.fillna('')
    return df

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload-kronos', methods=['POST'])
def upload_kronos():
    # ... (esta função permanece inalterada)
    print("Iniciando upload de Kronos...")
    try:
        if 'file' not in request.files:
            return jsonify({'message': 'Nenhum arquivo enviado.'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'message': 'Nenhum arquivo selecionado.'}), 400
        if file:
            filename = "Base.xlsx"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            print("Arquivo Kronos salvo. Lendo dados para visualização.")
            df = pd.read_excel(file_path)
            df = sanitize_dataframe(df)
            data = df.to_dict('records')
            return jsonify({
                'message': 'Arquivo Kronos processado com sucesso.',
                'data': data,
                'headers': list(df.columns)
            })
    except Exception as e:
        print(f"Erro no upload de Kronos: {e}")
        return jsonify({'message': f'Ocorreu um erro no servidor: {e}'}), 500

@app.route('/upload-sales', methods=['POST'])
def upload_sales():
    # ... (esta função permanece inalterada)
    print("Iniciando upload de Sales...")
    try:
        if 'file' not in request.files:
            return jsonify({'message': 'Nenhum arquivo enviado.'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'message': 'Nenhum arquivo selecionado.'}), 400
        filename = "Sales.xlsx"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        print("Arquivo Sales salvo. Lendo dados para visualização.")
        df = pd.read_excel(file_path, engine='openpyxl')
        df = df.drop(columns=['Unnamed: 0'], errors='ignore')
        df = sanitize_dataframe(df)
        data = df.to_dict('records')
        headers = list(df.columns)
        return jsonify({
            'message': 'Arquivo Sales processado com sucesso.',
            'data': data,
            'headers': headers
        })
    except Exception as e:
        print(f"**ERRO NO SERVIDOR:** {e}")
        return jsonify({'message': f'Ocorreu um erro ao processar o arquivo: {e}'}), 500

# AJUSTE: Rota para lançar a reposição de Sales
@app.route('/lancar-sales', methods=['POST'])
def lancar_sales():
    print("Botão 'Lançar Reposição Sales' foi clicado. Iniciando automação...")
    try:
        # Define o caminho completo para o arquivo que foi enviado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Sales.xlsx')
        
        # Verifica se o arquivo existe antes de iniciar
        if not os.path.exists(file_path):
            return jsonify({'message': 'Erro: Arquivo Sales.xlsx não encontrado. Faça o upload primeiro.'}), 404
            
        # Executa o script Sales.py em segundo plano, passando o caminho do arquivo como argumento
        # sys.executable garante que o script use o mesmo interpretador Python que a aplicação Flask
        subprocess.Popen([sys.executable, 'Sales.py', file_path])
        
        return jsonify({'message': 'Automação de Sales iniciada com sucesso!'})
    except Exception as e:
        print(f"Erro ao lançar reposição Sales: {e}")
        return jsonify({'message': f'Erro ao iniciar a automação: {e}'}), 500

# AJUSTE: Rota para lançar a reposição de Kronos
@app.route('/lancar-kronos', methods=['POST'])
def lancar_kronos():
    print("Botão 'Lançar Reposição Kronos' foi clicado. Iniciando automação...")
    try:
        # Define o caminho completo para o arquivo que foi enviado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Base.xlsx')
        
        # Verifica se o arquivo existe antes de iniciar
        if not os.path.exists(file_path):
            return jsonify({'message': 'Erro: Arquivo Base.xlsx não encontrado. Faça o upload primeiro.'}), 404
        
        # Executa o script Kronos.py em segundo plano
        subprocess.Popen([sys.executable, 'Kronos.py', file_path])
        
        return jsonify({'message': 'Automação de Kronos iniciada com sucesso!'})
    except Exception as e:
        print(f"Erro ao lançar reposição Kronos: {e}")
        return jsonify({'message': f'Erro ao iniciar a automação: {e}'}), 500

if __name__ == '__main__':
    app.run(debug=True)