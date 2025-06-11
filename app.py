from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os
import json
import subprocess
import tempfile
from pathlib import Path

app = Flask(__name__)
app.secret_key = 'chave_super_secreta'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Diretórios do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
password_dir = os.path.join(BASE_DIR, "password")
log_dir = os.path.join(BASE_DIR, "static", "logs")
arquivos_dir = os.path.join(BASE_DIR, "static", "arquivos")
os.makedirs(password_dir, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)
os.makedirs(arquivos_dir, exist_ok=True)

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

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

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

@app.route('/memoriais-descritivos', methods=['GET', 'POST'])
def memoriais_descritivos():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = zip_download = log_relativo = None

    if request.method == 'POST':
        diretorio = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
        cidade = request.form['cidade']
        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']

        os.makedirs(diretorio, exist_ok=True)

        caminho_excel = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_excel.filename)
        caminho_dxf = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_dxf.filename)
        arquivo_excel.save(caminho_excel)
        arquivo_dxf.save(caminho_dxf)

        log_filename = datetime.now().strftime("log_%Y%m%d_%H%M%S.txt")
        log_path = os.path.join(log_dir, log_filename)

        try:
            with open(log_path, 'w', encoding='utf-8') as log_file:
                try:
                    log_file.write("✅ Simulação de execução!\n")
                    log_file.write(f"Cidade: {cidade}\n")
                    log_file.write(f"Excel: {caminho_excel}\n")
                    log_file.write(f"DXF: {caminho_dxf}\n")

                    # Criar arquivo simulado visível
                    simulacao_path = os.path.join(arquivos_dir, "simulacao.txt")
                    with open(simulacao_path, "w", encoding="utf-8") as f:
                        f.write("📄 Arquivo gerado via simulação no app.py\n")

                    log_file.write("📦 Arquivo simulacao.txt criado com sucesso.\n")
                    resultado = "✅ Execução simulada com sucesso! Arquivo gerado."
                    log_relativo = f"logs/{log_filename}"

    except Exception as e:
        erro_execucao = f"❌ Erro na simulação:<br><pre>{type(e).__name__}: {e}</pre>"


            if processo.returncode == 0:
                resultado = "✅ Processamento concluído com sucesso!"
                log_relativo = f"logs/{log_filename}"
            else:
                with open(log_path, 'r', encoding='utf-8') as log_file:
                    erro_execucao = f"❌ Erro na execução:<br><pre>{log_file.read()}</pre>"

        except Exception as e:
            try:
                with open(log_path, 'r', encoding='utf-8') as log_file:
                    erro_execucao = f"❌ Erro inesperado:<br><pre>{log_file.read()}</pre>"
            except Exception as leitura_erro:
                erro_execucao = f"❌ Erro inesperado e falha ao ler log: {leitura_erro}"

        finally:
            os.remove(caminho_excel)
            os.remove(caminho_dxf)

        try:
            arquivos_zip = [f for f in os.listdir(diretorio) if f.lower().endswith('.zip')]
            if arquivos_zip:
                arquivos_zip.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)), reverse=True)
                zip_download = arquivos_zip[0]
        except Exception as e:
            print(f"⚠️ Erro ao localizar arquivo ZIP para download: {e}")

    return render_template("formulario_DECOPA.html", resultado=resultado, erro=erro_execucao, zip_download=zip_download, log_path=log_relativo)

@app.route("/arquivos-gerados")
def listar_arquivos_gerados():
    if not os.path.exists(arquivos_dir):
        return "<h3>⚠️ Nenhum diretório 'static/arquivos' encontrado.</h3>"
    arquivos = list(Path(arquivos_dir).glob("*.*"))
    if not arquivos:
        return "<h3>📭 Nenhum arquivo foi gerado ainda.</h3>"
    links_html = "".join(f'<li><a href="/static/arquivos/{a.name}" download>{a.name}</a></li>' for a in arquivos)
    return f"<h2>📂 Arquivos Gerados:</h2><ul>{links_html}</ul><p><a href='/'>🔙 Voltar para o início</a></p>"

@app.route('/download-zip/<filename>')
def download_zip(filename):
    caminho_zip = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', filename)
    if os.path.exists(caminho_zip):
        return send_file(caminho_zip, as_attachment=True)
    else:
        return f"Arquivo {filename} não encontrado.", 404

@app.route("/downloads")
def listar_arquivos():
    os.makedirs(arquivos_dir, exist_ok=True)
    arquivos = os.listdir(arquivos_dir)
    return render_template("listar_arquivos.html", arquivos=arquivos)

@app.route("/download/<nome_arquivo>")
def download_arquivo(nome_arquivo):
    return send_from_directory(arquivos_dir, nome_arquivo, as_attachment=True)

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

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
