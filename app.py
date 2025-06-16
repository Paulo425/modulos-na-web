from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os
import json
import subprocess
import tempfile
from pathlib import Path
from subprocess import Popen, PIPE
import shutil
import os
from werkzeug.security import generate_password_hash, check_password_hash

from usuarios_mysql import (
    salvar_usuario_mysql,
    buscar_usuario_mysql,
    aprovar_usuario_mysql,
    excluir_usuario_mysql,
    listar_pendentes_mysql,
    listar_usuarios_mysql,
    atualizar_senha_mysql
)




BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)  # ✅ Cria pasta em tempo de execução

def listar_usuarios_mysql():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT usuario FROM usuarios ORDER BY usuario")
    resultado = [row[0] for row in cursor.fetchall()]
    conn.close()
    return resultado

def atualizar_senha_mysql(usuario, nova_senha_hash):
    conn = get_db_connection()
    cursor = conn.cursor()
    sql = "UPDATE usuarios SET senha_hash = %s WHERE usuario = %s"
    cursor.execute(sql, (nova_senha_hash, usuario))
    conn.commit()
    conn.close()

app = Flask(__name__)
app.secret_key = 'chave_super_secreta'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Diretórios do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, "static", "logs")
arquivos_dir = os.path.join(BASE_DIR, "static", "arquivos")
os.makedirs(log_dir, exist_ok=True)
os.makedirs(arquivos_dir, exist_ok=True)

@app.context_processor
def inject_pendentes_count():
    if session.get('usuario') == 'admin':
        try:
            return dict(pendentes_count=len(listar_pendentes_mysql()))
        except:
            return dict(pendentes_count=0)
    return dict(pendentes_count=0)


@app.route('/')
def home():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    pendentes_count = 0
    if session.get('usuario') == 'admin':
        pendentes_count = len(listar_pendentes_mysql())

    return render_template('index.html', pendentes_count=pendentes_count)




@app.route('/login', methods=['GET', 'POST'])
def login():
    erro = None
    debug = None

    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']

        try:
            dados = buscar_usuario_mysql(usuario)

            if not dados:
                erro = "Usuário ou senha inválidos."
            else:
                senha_hash = dados.get("senha_hash")
                aprovado = dados.get("aprovado", True)

                if not aprovado:
                    erro = "Conta ainda não aprovada. Aguarde a autorização do administrador."
                elif check_password_hash(senha_hash, senha):
                    session['usuario'] = usuario
                    return redirect(url_for('home'))
                else:
                    erro = "Usuário ou senha inválidos."

        except Exception as e:
            erro = "Erro ao processar login."
            debug = f"{type(e).__name__}: {str(e)}"

    return render_template('login.html', erro=erro, debug=debug)


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

        existente = buscar_usuario_mysql(novo_usuario)
        if existente:
            erro = f"Usuário '{novo_usuario}' já existe."
        else:
            senha_hash = generate_password_hash(nova_senha)
            salvar_usuario_mysql(novo_usuario, senha_hash, nivel='tecnico', aprovado=True)
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
            existente = buscar_usuario_mysql(usuario)
            if existente:
                excluir_usuario_mysql(usuario)
                mensagem = f"Usuário '{usuario}' excluído com sucesso."
            else:
                erro = f"Usuário '{usuario}' não encontrado."

    usuarios = listar_usuarios_mysql()
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

        # Corrigido para salvar o log na pasta pública correta
        log_filename = datetime.now().strftime("log_%Y%m%d_%H%M%S.log")

        # Garante que o log seja salvo em /static/logs no diretório RAIZ do projeto
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)

        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"

        # DEBUG opcional
        print(f"🧾 Salvando LOG em: {log_path}")

        
        try:
            processo = Popen(
                ["python", os.path.join(BASE_DIR, "executaveis", "main.py"),
                 "--diretorio", diretorio,
                 "--cidade", cidade,
                 "--excel", caminho_excel,
                 "--dxf", caminho_dxf],
                stdout=PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            log_lines = []
            with open(log_path, 'w', encoding='utf-8') as log_file:
                for linha in processo.stdout:
                    if len(log_lines) < 100:
                        log_file.write(linha)
                        log_lines.append(linha)
                    print("🖨️", linha.strip())

            processo.wait()

            if processo.returncode == 0:
                resultado = "✅ Processamento concluído com sucesso!"
            else:
                erro_execucao = f"❌ Erro na execução:<br><pre>{''.join(log_lines)}</pre>"

        except Exception as e:
            erro_execucao = f"❌ Erro inesperado:<br><pre>{type(e).__name__}: {str(e)}</pre>"


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

@app.route('/download/<filename>')
def download_zip(filename):
    caminho = os.path.join(BASE_DIR, 'static', 'arquivos')
    return send_from_directory(caminho, filename, as_attachment=True)


@app.route('/registrar', methods=['GET', 'POST'])
def registrar():
    mensagem = erro = None
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']

        # Verifica se o usuário já existe no banco
        existente = buscar_usuario_mysql(usuario)
        if existente:
            erro = "Usuário já existe ou está aguardando aprovação."
        else:
            senha_hash = generate_password_hash(senha)
            salvar_usuario_mysql(usuario, senha_hash, nivel="tecnico", aprovado=False)
            mensagem = "Conta criada com sucesso! Aguarde autorização do administrador."

    return render_template('registrar.html', mensagem=mensagem, erro=erro)


@app.route('/pendentes', methods=['GET', 'POST'])
def pendentes():
    if session.get('usuario') != 'admin':
        return redirect(url_for('login'))

    if request.method == 'POST':
        aprovados = request.form.getlist('aprovar')
        for usuario in aprovados:
            aprovar_usuario_mysql(usuario)
        return redirect(url_for('pendentes'))

    usuarios_pendentes = listar_pendentes_mysql()
    return render_template('pendentes.html', pendentes=usuarios_pendentes)



@app.route('/alterar-senha', methods=['GET', 'POST'])
def alterar_senha():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mensagem = erro = None

    if request.method == 'POST':
        atual = request.form['senha_atual']
        nova = request.form['nova_senha']
        usuario = session['usuario']

        dados = buscar_usuario_mysql(usuario)

        if not dados:
            erro = "Usuário não encontrado."
        elif check_password_hash(dados['senha_hash'], atual):
            nova_hash = generate_password_hash(nova)
            atualizar_senha_mysql(usuario, nova_hash)
            mensagem = "Senha alterada com sucesso!"
        else:
            erro = "Senha atual incorreta."

    return render_template('alterar_senha.html', mensagem=mensagem, erro=erro)


@app.route("/downloads")
def listar_arquivos():
    os.makedirs(arquivos_dir, exist_ok=True)
    arquivos = os.listdir(arquivos_dir)
    return render_template("listar_arquivos.html", arquivos=arquivos)

@app.route("/download/<nome_arquivo>")
def download_arquivo(nome_arquivo):
    return send_from_directory(arquivos_dir, nome_arquivo, as_attachment=True)

@app.route('/memorial_azimute_az', methods=['GET', 'POST'])
def gerar_memorial_azimute_az():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        diretorio_tmp = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
        os.makedirs(diretorio_tmp, exist_ok=True)

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']
        caminho_excel = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_excel.filename)
        caminho_dxf = os.path.join(app.config['UPLOAD_FOLDER'], arquivo_dxf.filename)
        arquivo_excel.save(caminho_excel)
        arquivo_dxf.save(caminho_dxf)

        log_filename = datetime.now().strftime("log_AZIMUTEAZ_%Y%m%d_%H%M%S.log")
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)
        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"

        try:
            processo = Popen(
                ["python", os.path.join(BASE_DIR, "executaveis_azimute_az", "main.py"),
                 cidade, caminho_excel, caminho_dxf],
                stdout=PIPE, stderr=subprocess.STDOUT, text=True
            )

            log_lines = []
            with open(log_path, 'w', encoding='utf-8') as log_file:
                for linha in processo.stdout:
                    log_file.write(linha)
                    if len(log_lines) < 500:
                        log_lines.append(linha)
                    print("🖨️", linha.strip())

            processo.wait()

            if processo.returncode == 0:
                resultado = "✅ Processamento concluído com sucesso!"
            else:
                erro_execucao = f"❌ Erro na execução:<br><pre>{''.join(log_lines)}</pre>"

        except Exception as e:
            erro_execucao = f"❌ Erro inesperado:<br><pre>{type(e).__name__}: {str(e)}</pre>"

        finally:
            os.remove(caminho_excel)
            os.remove(caminho_dxf)

        # 🔍 Verificação do ZIP após o processamento
        try:
            zip_dir = os.path.join(BASE_DIR, 'static', 'arquivos')
            arquivos_zip = [f for f in os.listdir(zip_dir) if f.lower().endswith('.zip')]
            if arquivos_zip:
                arquivos_zip.sort(key=lambda x: os.path.getmtime(os.path.join(zip_dir, x)), reverse=True)
                zip_download = arquivos_zip[0]
                print(f"✅ ZIP disponível para download: {zip_download}")
            else:
                print("⚠️ Nenhum ZIP encontrado no diretório público.")
        except Exception as e:
            print(f"❌ Erro ao verificar ZIP: {e}")
            zip_download = None
    return render_template("formulario_AZIMUTE_AZ.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)





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


@app.route('/debug_rotas')
def debug_rotas():
    from flask import Response
    rotas = [str(rule) for rule in app.url_map.iter_rules()]
    return Response('<br>'.join(rotas), mimetype='text/html')


