# app.py
from flask import Flask, render_template, request, redirect, url_for, session
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os
import tempfile
import json

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.secret_key = 'sua_chave_secreta_aqui'

# Caminhos compatíveis com Render (Linux) e Windows
base_path = os.path.dirname(os.path.abspath(__file__))
password_dir = os.path.join(base_path, "password")
dir_log_projeto = os.path.join(base_path, "Log")
os.makedirs(password_dir, exist_ok=True)
os.makedirs(dir_log_projeto, exist_ok=True)

admin_path = os.path.join(password_dir, "admin.json")
admin_data = {
    "usuario": "admin",
    "senha_hash": "pbkdf2:sha256:600000$MUVq0v03EaDI0pRV$bb9a268b95b1baf86c6277ef1cdadcc3c7c443b2bc9608c48a574e8dc1de76c0"
}
with open(admin_path, 'w', encoding='utf-8') as f:
    json.dump(admin_data, f, indent=2)
print("🛠️ admin.json recriado com senha 1234")


        
def carregar_usuarios():
    usuarios = {}
    for nome_arquivo in os.listdir(password_dir):
        if nome_arquivo.endswith('.json'):
            caminho = os.path.join(password_dir, nome_arquivo)
            with open(caminho, 'r', encoding='utf-8') as f:
                dados = json.load(f)
                usuarios[dados['usuario']] = dados['senha_hash']
    return usuarios

@app.route('/login', methods=['GET', 'POST'])
def login():
    erro = None
    usuarios = carregar_usuarios()
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']
        if usuario in usuarios and check_password_hash(usuarios[usuario], senha):
            session['usuario'] = usuario
            return redirect(url_for('memoriais_descritivos'))
        else:
            erro = 'Usuário ou senha inválidos!'
    return render_template('login.html', erro=erro)

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/memoriais-descritivos')
def memoriais_descritivos():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    return render_template('memoriais_descritivos.html')

@app.route('/executar-decopa', methods=['GET', 'POST'])
def executar_decopa():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    erro_execucao = None
    resultado = None

    if request.method == 'POST':
        diretorio = request.form['diretorio'].replace('"', '')
        cidade = request.form['cidade']

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']

        caminho_excel = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_excel.filename)
        caminho_dxf = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_dxf.filename)

        arquivo_excel.save(caminho_excel)
        arquivo_dxf.save(caminho_dxf)

        # Substituindo a execução do .exe por mensagem informativa
        resultado = "Executável DECOPA não está disponível nesta versão online (Render)."

        # Limpar arquivos após a simulação
        if os.path.exists(caminho_excel):
            os.remove(caminho_excel)
        if os.path.exists(caminho_dxf):
            os.remove(caminho_dxf)

    return render_template('formulario_DECOPA.html', resultado=resultado, erro=erro_execucao)

@app.route('/criar-usuario', methods=['GET', 'POST'])
def criar_usuario():
    if session.get('usuario') != 'admin':
        return redirect(url_for('login'))

    mensagem = None
    erro = None

    if request.method == 'POST':
        novo_usuario = request.form['usuario']
        nova_senha = request.form['senha']

        caminho_arquivo = os.path.join(password_dir, f"{novo_usuario}.json")

        if os.path.exists(caminho_arquivo):
            erro = f"O usuário '{novo_usuario}' já existe."
        else:
            dados = {
                "usuario": novo_usuario,
                "senha_hash": generate_password_hash(nova_senha)
            }
            with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                json.dump(dados, f, indent=2)
            mensagem = f"Usuário '{novo_usuario}' criado com sucesso."

    return render_template('criar_usuario.html', mensagem=mensagem, erro=erro)

@app.route('/excluir-usuario', methods=['GET', 'POST'])
def excluir_usuario():
    if session.get('usuario') != 'admin':
        return redirect(url_for('login'))

    mensagem = None
    erro = None

    if request.method == 'POST':
        usuario_para_excluir = request.form['usuario']
        if usuario_para_excluir == 'admin':
            erro = "Não é permitido excluir o usuário 'admin'."
        else:
            caminho = os.path.join(password_dir, f"{usuario_para_excluir}.json")
            if os.path.exists(caminho):
                os.remove(caminho)
                mensagem = f"Usuário '{usuario_para_excluir}' excluído com sucesso."
            else:
                erro = f"O usuário '{usuario_para_excluir}' não foi encontrado."

    usuarios_atuais = [f[:-5] for f in os.listdir(password_dir) if f.endswith('.json')]

    return render_template('excluir_usuario.html', usuarios=usuarios_atuais, mensagem=mensagem, erro=erro)

if __name__ == '__main__':
    import os
    port = os.environ['PORT']
    print(f"🔵 Porta definida pela Render: {port}")
    app.run(host='0.0.0.0', port=int(port))

#atualiza



