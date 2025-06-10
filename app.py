from flask import Flask, render_template, request, redirect, url_for, session
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os
import json
import subprocess
import tempfile

app = Flask(__name__)
app.secret_key = 'chave_super_secreta'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Diretórios do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
password_dir = os.path.join(BASE_DIR, "password")
log_dir = os.path.join(BASE_DIR, "Log")
os.makedirs(password_dir, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)

# Criação do admin.json
admin_path = os.path.join(password_dir, "admin.json")
if not os.path.exists(admin_path):
    admin_user = {
        "usuario": "admin",
        "senha_hash": generate_password_hash("1234")
    }
    with open(admin_path, 'w', encoding='utf-8') as f:
        json.dump(admin_user, f, indent=2)
    print("✅ admin.json criado com senha 1234")
else:
    print("🔹 admin.json já existe")

# Carregar usuários
def carregar_usuarios():
    usuarios = {}
    for arquivo in os.listdir(password_dir):
        if arquivo.endswith('.json'):
            with open(os.path.join(password_dir, arquivo), 'r', encoding='utf-8') as f:
                dados = json.load(f)
                usuarios[dados['usuario']] = dados['senha_hash']
    return usuarios

@app.route('/')
def home():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    return render_template('index.html')

# Login
@app.route('/login', methods=['GET', 'POST'])
def login():
    erro = None
    debug = None

    try:
        usuarios = carregar_usuarios()
    except Exception as e:
        erro = "Erro ao carregar usuários!"
        debug = f"{type(e).__name__}: {str(e)}"
        return render_template('login.html', erro=erro, debug=debug)

    if request.method == 'POST':
        try:
            usuario = request.form['usuario']
            senha = request.form['senha']
            if usuario in usuarios and check_password_hash(usuarios[usuario], senha):
                session['usuario'] = usuario
                return redirect(url_for('home'))
            else:
                erro = "Usuário ou senha inválidos!"
        except Exception as e:
            erro = "Erro interno no processamento do login."
            debug = f"{type(e).__name__}: {str(e)}"
            return render_template('login.html', erro=erro, debug=debug)

    return render_template('login.html', erro=erro)


# Logout
@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

# Criar usuário
@app.route('/criar-usuario', methods=['GET', 'POST'])
def criar_usuario():
    if session.get('usuario') != 'admin':
        return redirect(url_for('login'))
    mensagem = erro = None
    if request.method == 'POST':
        novo_usuario = request.form['usuario']
        nova_senha = request.form['senha']
        caminho = os.path.join(password_dir, f"{novo_usuario}.json")
        if os.path.exists(caminho):
            erro = f"Usuário '{novo_usuario}' já existe."
        else:
            dados = {
                "usuario": novo_usuario,
                "senha_hash": generate_password_hash(nova_senha)
            }
            with open(caminho, 'w', encoding='utf-8') as f:
                json.dump(dados, f, indent=2)
            mensagem = f"Usuário '{novo_usuario}' criado com sucesso!"
    return render_template('criar_usuario.html', mensagem=mensagem, erro=erro)

# Excluir usuário
@app.route('/excluir-usuario', methods=['GET', 'POST'])
def excluir_usuario():
    if session.get('usuario') != 'admin':
        return redirect(url_for('login'))
    mensagem = erro = None
    if request.method == 'POST':
        usuario = request.form['usuario']
        if usuario == 'admin':
            erro = "Não é permitido excluir o usuário 'admin'."
        else:
            caminho = os.path.join(password_dir, f"{usuario}.json")
            if os.path.exists(caminho):
                os.remove(caminho)
                mensagem = f"Usuário '{usuario}' excluído com sucesso."
            else:
                erro = f"Usuário '{usuario}' não encontrado."
    usuarios = [f[:-5] for f in os.listdir(password_dir) if f.endswith('.json')]
    return render_template('excluir_usuario.html', usuarios=usuarios, mensagem=mensagem, erro=erro)

# Módulo: Memoriais Descritivos (DECOPA)
@app.route('/memoriais-descritivos', methods=['GET', 'POST'])
def memoriais_descritivos():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = None

    if request.method == 'POST':
        diretorio = request.form['diretorio'].replace('"', '')
        cidade = request.form['cidade']
        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']

        caminho_excel = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_excel.filename)
        caminho_dxf = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_dxf.filename)
        arquivo_excel.save(caminho_excel)
        arquivo_dxf.save(caminho_dxf)

        log_filename = datetime.now().strftime("log_%Y%m%d_%H%M%S.txt")
        log_path = os.path.join(log_dir, log_filename)

        try:
            with open(log_path, 'w', encoding='utf-8') as log_file:
                processo = subprocess.run(
                    [
                        "python", os.path.join(BASE_DIR, "executaveis", "main.py"),
                        "--diretorio", diretorio,
                        "--cidade", cidade,
                        "--excel", caminho_excel,
                        "--dxf", caminho_dxf
                    ],
                    stdout=log_file,
                    stderr=subprocess.STDOUT,
                    encoding='utf-8'
                )
            if processo.returncode == 0:
                resultado = f"✅ Processamento concluído! Arquivos salvos em {diretorio}."
            else:
                with open(log_path, 'r', encoding='utf-8') as log_file:
                    log_conteudo = log_file.read()
                erro_execucao = f"❌ Erro na execução:<br><pre>{log_conteudo}</pre>"
        except Exception as e:
            try:
                with open(log_path, 'r', encoding='utf-8') as log_file:
                    log_conteudo = log_file.read()
                erro_execucao = f"❌ Erro na execução:<br><pre>{log_conteudo}</pre>"
            except Exception as leitura_erro:
                erro_execucao = f"❌ Erro na execução e também falhou ao ler o log: {leitura_erro}"
        finally:
            os.remove(caminho_excel)
            os.remove(caminho_dxf)

    return render_template('formulario_DECOPA.html', resultado=resultado, erro=erro_execucao)


# Módulos futuros
@app.route('/memoriais-azimute-az')
def memoriais_azimute_az():
    return render_template('em_breve.html', titulo="MEMORIAIS_AZIMUTE_AZ")

@app.route('/memoriais-azimute-p1-p2')
def memoriais_azimute_p1_p2():
    return render_template('em_breve.html', titulo="MEMORIAIS_AZIMUTE_P1_P2")

@app.route('/memoriais-angulos-internos-az')
def memoriais_angulos_internos_az():
    return render_template('em_breve.html', titulo="MEMORIAIS_ANGULOS_INTERNOS_AZ")

@app.route('/memoriais-angulos-internos-p1-p2')
def memoriais_angulos_internos_p1_p2():
    return render_template('em_breve.html', titulo="MEMORIAIS_ANGULOS_INTERNOS_P1_P2")

# Roda o servidor local
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
