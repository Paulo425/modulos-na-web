from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from subprocess import Popen, DEVNULL, STDOUT
from subprocess import Popen, PIPE, STDOUT
import os
import json
import subprocess
import tempfile
from pathlib import Path
from subprocess import Popen, PIPE
import shutil
import os
from werkzeug.security import generate_password_hash, check_password_hash
import traceback
import sys 
import fitz  # PyMuPDF
from flask import flash
import zipfile
from flask import send_file



from usuarios_mysql import (
    salvar_usuario_mysql,
    buscar_usuario_mysql,
    aprovar_usuario_mysql,
    excluir_usuario_mysql,
    listar_pendentes_mysql,
    listar_usuarios_mysql,
    atualizar_senha_mysql
)

import logging
import sys
import uuid
import logging, traceback
from datetime import datetime
from pdf2image import convert_from_bytes
import io
from PIL import Image, UnidentifiedImageError
           


logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s: %(message)s',
    stream=sys.stdout
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)  # ✅ Cria pasta em tempo de execução


app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'),
                       static_folder=os.path.join(BASE_DIR, 'static'))

app.secret_key = 'chave_super_secreta'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Diretórios do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, "static", "logs")
arquivos_dir = os.path.join(BASE_DIR, "static", "arquivos")
os.makedirs(log_dir, exist_ok=True)
os.makedirs(arquivos_dir, exist_ok=True)

# @app.context_processor
# def inject_pendentes_count():
#     if session.get('usuario') == 'admin':
#         try:
#             return dict(pendentes_count=len(listar_pendentes_mysql()))
#         except:
#             return dict(pendentes_count=0)
#     return dict(pendentes_count=0)
import uuid

def salvar_com_nome_unico(arquivo, destino_base):
    """
    Salva o arquivo com um nome único no destino_base.
    Retorna o caminho completo salvo.
    """
    uuid_curtinho = uuid.uuid4().hex[:8]
    nome_unico = f"{uuid_curtinho}_{arquivo.filename}"
    caminho_completo = os.path.join(destino_base, nome_unico)
    arquivo.save(caminho_completo)
    return caminho_completo


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
                print("🔴 Usuário não encontrado no banco.")
            else:
                senha_hash = dados.get("senha_hash")
                aprovado = dados.get("aprovado", True)

                logging.info("🔍 DEBUG LOGIN:")
                logging.info(f"Usuário digitado: {usuario}")
                logging.info(f"Senha digitada : {senha}")
                logging.info(f"Hash no banco   : {senha_hash}")
                logging.info(f"Aprovado        : {aprovado} ({type(aprovado)})")

                # Interpretação segura de 'aprovado'
                aprovado_bool = (
                    bool(aprovado) if isinstance(aprovado, bool)
                    else str(aprovado).strip().lower() in ['1', 'true', 'yes']
                )

                if not aprovado_bool:
                    erro = "Conta ainda não aprovada. Aguarde a autorização do administrador."
                    print("🔴 Conta não aprovada.")
                elif not senha_hash or not check_password_hash(senha_hash, senha):
                    erro = "Usuário ou senha inválidos."
                    print("🔴 Senha incorreta para esse hash.")
                else:
                    print("✅ Login autorizado. Redirecionando...")
                    session['usuario'] = usuario
                    return redirect(url_for('home'))

        except Exception as e:
            erro = "Erro ao processar login."
            debug = f"{type(e).__name__}: {str(e)}"
            print(f"❌ Erro durante login: {debug}")

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
        import uuid

        id_execucao = str(uuid.uuid4())[:8]
        diretorio = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)
        os.makedirs(diretorio, exist_ok=True)

        cidade = request.form['cidade']
        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']

        caminho_excel = salvar_com_nome_unico(arquivo_excel, app.config['UPLOAD_FOLDER'])
        caminho_dxf   = salvar_com_nome_unico(arquivo_dxf, app.config['UPLOAD_FOLDER'])

        # Log
        log_filename = datetime.now().strftime("log_%Y%m%d_%H%M%S.log")
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)
        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"
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
                    log_file.write(linha)
                    if len(log_lines) < 100:
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

        # Verifica ZIP e copia para static/arquivos
        try:
            static_zip_dir = os.path.join(BASE_DIR, 'static', 'arquivos')
            arquivos_zip = [f for f in os.listdir(static_zip_dir) if f.lower().endswith('.zip')]

            print("🧪 ZIPs encontrados:", arquivos_zip)
            logging.info(f"🧪 ZIPs encontrados: {arquivos_zip}")

            if arquivos_zip:
                arquivos_zip.sort(key=lambda x: os.path.getmtime(os.path.join(static_zip_dir, x)), reverse=True)
                zip_download = arquivos_zip[0]
                print(f"✅ ZIP para download: {zip_download}")
                logging.info(f"✅ ZIP para download: {zip_download}")
        except Exception as e:
            print(f"⚠️ Erro ao localizar/copiar ZIP: {e}")
            logging.error(f"⚠️ Erro ao localizar/copiar ZIP: {e}")

    return render_template("formulario_DECOPA.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)
#ATUALIZADO


@app.route("/arquivos-gerados")
def listar_arquivos_gerados():
    if not os.path.exists(arquivos_dir):
        return "<h3>⚠️ Nenhum diretório 'static/arquivos' encontrado.</h3>"
    arquivos = list(Path(arquivos_dir).glob("*.*"))
    if not arquivos:
        return "<h3>📭 Nenhum arquivo foi gerado ainda.</h3>"
    links_html = "".join(f'<li><a href="/static/arquivos/{a.name}" download>{a.name}</a></li>' for a in arquivos)
    return f"<h2>📂 Arquivos Gerados:</h2><ul>{links_html}</ul><p><a href='/'>🔙 Voltar para o início</a></p>"

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

# @app.route("/download/<nome_arquivo>")
# def download_arquivo(nome_arquivo):
#     return send_from_directory(arquivos_dir, nome_arquivo, as_attachment=True)

@app.route('/memorial_azimute_az', methods=['GET', 'POST'])
def gerar_memorial_azimute_az():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        id_execucao = str(uuid.uuid4())[:8]
        diretorio_tmp = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)
        os.makedirs(diretorio_tmp, exist_ok=True)

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']
        caminho_excel = salvar_com_nome_unico(arquivo_excel, app.config['UPLOAD_FOLDER'])
        caminho_dxf   = salvar_com_nome_unico(arquivo_dxf, app.config['UPLOAD_FOLDER'])


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
            #parent_dir = os.path.dirname(diretorio)  # sobe de tmp/CONCLUIDO/uuid para tmp/CONCLUIDO
            # Verifica ZIP na pasta correta (static/arquivos)
            static_zip_dir = os.path.join(BASE_DIR, 'static', 'arquivos')
            arquivos_zip = [f for f in os.listdir(static_zip_dir) if f.lower().endswith('.zip')]


            print("🧪 ZIPs disponíveis:", arquivos_zip)
            logging.info(f"🧪 ZIPs disponíveis: {arquivos_zip}")

            if arquivos_zip:
                arquivos_zip.sort(key=lambda x: os.path.getmtime(os.path.join(static_zip_dir, x)), reverse=True)
                zip_download = arquivos_zip[0]
                print(f"✅ ZIP disponível para download: {zip_download}")
                logging.info(f"✅ ZIP disponível para download: {zip_download}")

        except Exception as e:
            print(f"⚠️ Erro ao localizar/copiar ZIP: {e}")
            logging.error(f"⚠️ Erro ao localizar/copiar ZIP: {e}")

    print("DEBUG FLASK - zip_download final:", zip_download)
    

    return render_template("formulario_AZIMUTE_AZ.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)

@app.route('/download/<filename>')
def download_zip(filename):
    caminho = os.path.join(BASE_DIR, 'static', 'arquivos')
    return send_from_directory(caminho, filename, as_attachment=True)


@app.route('/memorial_azimute_jl', methods=['GET', 'POST'])
def memorial_azimute_jl():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = zip_download = log_relativo = None

    if request.method == 'POST':
        try:
            from executaveis.executar_memorial_azimute_jl import executar_memorial_jl
            import uuid, zipfile

            # 1. Inputs do formulário
            proprietario = request.form['proprietario']
            matricula = request.form['matricula']
            descricao = request.form['descricao']
            excel_file = request.files['excel_file']
            dxf_file = request.files['dxf_file']

            # 2. Preparar diretórios
            id_execucao = str(uuid.uuid4())[:8]
            pasta_execucao = f'memorial_jl_{id_execucao}'
            pasta_temp = os.path.join(BASE_DIR, 'static', 'arquivos', pasta_execucao)
            os.makedirs(pasta_temp, exist_ok=True)

            # 3. Salvar arquivos enviados
            excel_path = os.path.join(pasta_temp, 'confrontantes.xlsx')
            dxf_path = os.path.join(pasta_temp, 'original.dxf')
            excel_file.save(excel_path)
            dxf_file.save(dxf_path)

            # 4. Criar caminho do LOG no mesmo padrão do DECOPA
            log_filename = f"log_JL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
            os.makedirs(log_dir_absoluto, exist_ok=True)
            log_path = os.path.join(log_dir_absoluto, log_filename)
            log_relativo = f"logs/{log_filename}"

            # 5. Executar processo principal
            log_path_gerado, arquivos_gerados = executar_memorial_jl(
                proprietario=proprietario,
                matricula=matricula,
                descricao=descricao,
                caminho_salvar=pasta_temp,
                dxf_path=dxf_path,
                excel_path=excel_path,
                log_path=log_path  # passa explicitamente
            )

            # 6. Gerar ZIP com os arquivos de saída (sem o LOG)
            zip_name = f"memorial_{matricula}.zip"
            zip_path = os.path.join(BASE_DIR, 'static', 'arquivos', zip_name)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for arquivo in arquivos_gerados:

                    if os.path.exists(arquivo) and not arquivo.endswith('.log'):
                        zipf.write(arquivo, arcname=os.path.basename(arquivo))


            resultado = "✅ Processamento concluído com sucesso!"
            zip_download = zip_name

        except Exception as e:
            erro_execucao = f"❌ Erro na execução: {e}"

    return render_template("formulario_azimute_jl.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)




@app.route('/memorial_angulo_az', methods=['GET', 'POST'])
def gerar_memorial_angulo_az():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        id_execucao = str(uuid.uuid4())[:8]
        diretorio_tmp = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)

        os.makedirs(diretorio_tmp, exist_ok=True)

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']
        caminho_excel = salvar_com_nome_unico(arquivo_excel, app.config['UPLOAD_FOLDER'])
        caminho_dxf   = salvar_com_nome_unico(arquivo_dxf, app.config['UPLOAD_FOLDER'])


        log_filename = datetime.now().strftime("log_ANGULOAZ_%Y%m%d_%H%M%S.log")
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)
        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"

        try:
            processo = Popen(
                ["python", os.path.join(BASE_DIR, "executaveis_angulo_az", "main.py"),
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
    return render_template("formulario_ANGULO_AZ.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)




@app.route('/memorial_angulo_p1_p2', methods=['GET', 'POST'])
def gerar_memorial_angulo_p1_p2():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        id_execucao = str(uuid.uuid4())[:8]
        diretorio_tmp = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)
        os.makedirs(diretorio_tmp, exist_ok=True)

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']
        caminho_excel = salvar_com_nome_unico(arquivo_excel, app.config['UPLOAD_FOLDER'])
        caminho_dxf = salvar_com_nome_unico(arquivo_dxf, app.config['UPLOAD_FOLDER'])

        log_filename = datetime.now().strftime("log_ANGULO_P1_P2_%Y%m%d_%H%M%S.log")
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)
        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"

       

        try:
            processo = Popen(
                [sys.executable, os.path.join(BASE_DIR, "executaveis_angulo_p1_p2", "main.py"),
                cidade, caminho_excel, caminho_dxf],
                stdout=PIPE, stderr=STDOUT, text=True
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

        # 🔍 Verificação correta do ZIP após o processamento
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

    return render_template("formulario_angulo_p1_p2.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)

#ROTA AZIMUTE_P1_P2

@app.route('/memorial_azimute_p1_p2', methods=['GET', 'POST'])
def gerar_memorial_azimute_p1_p2():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        id_execucao = str(uuid.uuid4())[:8]
        diretorio_tmp = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)
        os.makedirs(diretorio_tmp, exist_ok=True)

        arquivo_excel = request.files['excel']
        arquivo_dxf = request.files['dxf']
        caminho_excel = salvar_com_nome_unico(arquivo_excel, app.config['UPLOAD_FOLDER'])
        caminho_dxf = salvar_com_nome_unico(arquivo_dxf, app.config['UPLOAD_FOLDER'])

        log_filename = datetime.now().strftime("log_AZIMUTE_P1_P2_%Y%m%d_%H%M%S.log")
        log_dir_absoluto = os.path.join(BASE_DIR, "static", "logs")
        os.makedirs(log_dir_absoluto, exist_ok=True)
        log_path = os.path.join(log_dir_absoluto, log_filename)
        log_relativo = f"static/logs/{log_filename}"

       

        try:
            processo = Popen(
                [sys.executable, os.path.join(BASE_DIR, "executaveis_azimute_p1_p2", "main.py"),
                cidade, caminho_excel, caminho_dxf],
                stdout=PIPE, stderr=STDOUT, text=True
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

        # 🔍 Verificação correta do ZIP após o processamento
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

    return render_template("formulario_azimute_p1_p2.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_relativo)

# ✅ ROTA PARA O MÓDULO DE AVALIAÇÕES DE IMOVEIS E PROPRIEDADES

from executaveis_avaliacao.main import gerar_relatorio_avaliacao_com_template

@app.route("/avaliacoes", methods=["GET", "POST"])
def gerar_avaliacao():


    # Defina o log_path como você fez:
    LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    log_path_relativo = f'logs/{os.path.basename(log_path)}'

    # Configuração avançada de logging (arquivo + console)
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Remove handlers existentes (evita duplicações)
    if logger.hasHandlers():
        logger.handlers.clear()

    # Handler para arquivo
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)

    # Handler para console (StreamHandler)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)

    # Formatação comum aos dois handlers
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Adiciona handlers ao logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    log_path_relativo = f'logs/{os.path.basename(log_path)}'
    logger.info(f"✅ Log criado em: {log_path_relativo}")
    
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = zip_download = log_relativo = None
    
    if request.method == "POST":
        logger.info("🔧 Início da execução do bloco POST em /avaliacoes")

        acao = request.form.get("acao", "").lower()
        try:
            from werkzeug.utils import secure_filename
            import uuid, zipfile

            # 1. Criação de diretório temporário para essa execução
            id_execucao = str(uuid.uuid4())[:8]
            pasta_execucao = f'avaliacao_{id_execucao}'
            pasta_temp = os.path.join(BASE_DIR, 'static', 'arquivos', pasta_execucao)
            os.makedirs(pasta_temp, exist_ok=True)
                    
                      
            # 2. Salvar arquivos recebidos
            caminho_planilha = os.path.join(pasta_temp, "planilha.xlsx")
            request.files["planilha_excel"].save(caminho_planilha)
            logger.info(f"✅ Planilha salva: {caminho_planilha} - {'existe' if os.path.exists(caminho_planilha) else 'NÃO existe'}")

             

            

            def salvar_multiplos(nome_form, prefixo):
                arquivos = request.files.getlist(nome_form)
                todos_grupos = []

                for i, arq in enumerate(arquivos):
                    if arq and arq.filename:
                        extensao = arq.filename.rsplit('.', 1)[-1].lower()
                        grupo_imagens = []

                        dados_arquivo = arq.read()

                        if extensao == "pdf":
                            nome_pdf_temporario = os.path.join(pasta_temp, f"{prefixo}_{i}.pdf")
                            with open(nome_pdf_temporario, "wb") as f:
                                f.write(dados_arquivo)

                            pdf = fitz.open(nome_pdf_temporario)
                            for p in range(pdf.page_count):
                                pix = pdf.load_page(p).get_pixmap(dpi=200)
                                nome_img = f"{prefixo}_{i}_{p}.png"
                                caminho_img = os.path.join(pasta_temp, nome_img)
                                pix.save(caminho_img)
                                grupo_imagens.append(caminho_img)
                                logger.info(f"✅ Página {p+1}/{pdf.page_count} salva: {caminho_img}")
                            pdf.close()
                        else:
                            try:
                                imagem = Image.open(io.BytesIO(dados_arquivo))
                                imagem.thumbnail((1024, 1024))
                                nome_img = secure_filename(f"{prefixo}_{i}.png")
                                caminho_img = os.path.join(pasta_temp, nome_img)
                                imagem.save(caminho_img, optimize=True, quality=70)
                                grupo_imagens.append(caminho_img)
                                logger.info(f"✅ Imagem salva: {caminho_img}")
                            except UnidentifiedImageError:
                                logger.error(f"❌ Arquivo inválido: {arq.filename}")
                                continue

                        if grupo_imagens:
                            todos_grupos.append(grupo_imagens)

                return todos_grupos




            fotos_imovel = salvar_multiplos("fotos_imovel", "foto_imovel")
            fotos_adicionais = salvar_multiplos("fotos_imovel_adicionais", "doc_adicional")
            fotos_proprietario = salvar_multiplos("doc_proprietario", "doc_proprietario")
            fotos_planta = salvar_multiplos("doc_planta", "planta")

            caminho_logo = ""
            logo = request.files.get("arquivo_logo")
            if logo and logo.filename:
                caminho_logo = os.path.join(pasta_temp, "logo.png")
                logo.save(caminho_logo)
                logger.info(f"✅ Logo salvo: {caminho_logo} - {'existe' if os.path.exists(caminho_logo) else 'NÃO existe'}")
            # 3. Inputs simples
            f = request.form
            def chk(nome): return f.get(nome, "").lower() == "sim"

            restricoes = []
            i = 1
            while f.get(f"tipo_restricao_{i}"):
                area = float(f.get(f"area_restricao_{i}", "0").replace(",", ".") or "0")
                perc = float(f.get(f"depreciacao_restricao_{i}", "0").replace(",", ".") or "0")
                restricoes.append({
                    "tipo": f.get(f"tipo_restricao_{i}"),
                    "area": area,
                    "percentualDepreciacao": perc,
                    "fator": (100.0 - perc) / 100.0
                })
                i += 1
            cidade = f.get("cidade", "").strip()

            fatores_usuario = {
                "nomeSolicitante": f.get("nome_solicitante"),
                "avaliadorNome": f.get("nome_avaliador"),
                "avaliadorRegistro": f.get("registro_avaliador"),
                "tipoImovel": f.get("tipo_imovel_escolhido"),
                "nomeProprietario": f.get("nome_proprietario"),
                "telefoneProprietario": f.get("telefone_proprietario") if chk("incluir_tel") else "Não Informado",
                "emailProprietario": f.get("email_proprietario") if chk("incluir_mail") else "Não Informado",
                "documentacaoImovel": f"Matrícula n° {f.get('num_doc')}" if f.get("num_doc") else "Documentação não informada",
                "nomeCartorio": f.get("nome_cartorio"),
                "nomeComarca": f.get("nome_comarca"),
                "enderecoCompleto": f.get("endereco_imovel"),
                "finalidade_descricao": f.get("finalidade_descricao") or f.get("finalidade_lido", ""),
                "area": chk("usar_fator_area"),
                "oferta": chk("usar_fator_oferta"),
                "aproveitamento": chk("usar_fator_aproveitamento"),
                "localizacao_mesma_regiao": chk("localizacao_mesma_regiao"),
                "topografia": chk("usar_fator_topografia"),
                "pedologia": chk("usar_fator_pedologia"),
                "pavimentacao": chk("usar_fator_pavimentacao"),
                "esquina": chk("usar_fator_esquina"),
                "acessibilidade": chk("usar_fator_acessibilidade"),
                "estrutura_escolha": f.get("estrutura_escolha", "").upper(),
                "conduta_escolha": f.get("conduta_escolha", "").upper(),
                "desempenho_escolha": f.get("desempenho_escolha", "").upper(),
                "caminhoLogo": caminho_logo,
                "restricoes": restricoes,
                "cidade": f.get("cidade", "").strip()


            }

            try:
                area_parcial = float(f.get("area_parcial", "0").replace(".", "").replace(",", "."))
            except:
                area_parcial = 0.0

            # 4. Geração do relatório
            nome_docx = "RELATORIO_AVALIACAO_COMPLETO.docx"
            caminho_docx = os.path.join(pasta_temp, nome_docx)

            from executaveis_avaliacao.main import (
                ler_planilha_excel, aplicar_chauvenet_e_filtrar,
                homogeneizar_amostras, gerar_grafico_aderencia_totais,
                gerar_grafico_dispersao_mediana
            )

            df_amostras, dados_imovel = ler_planilha_excel(caminho_planilha)
            logger.info(f"df_amostras.head():\n{df_amostras.head()}")
            logger.info(f"dados_imovel: {dados_imovel}")
            df_filtrado, idx_exc, amostras_exc, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_amostras)
            logger.info(f"df_filtrado.head():\n{df_filtrado.head()}")
            logger.info(f"Média: {media}, Mediana: {mediana}")
            homog = homogeneizar_amostras(df_filtrado, dados_imovel, fatores_usuario, "mercado")

            img1 = os.path.join(pasta_temp, "grafico_aderencia.png")
            img2 = os.path.join(pasta_temp, "grafico_dispersao.png")
            gerar_grafico_aderencia_totais(df_filtrado, homog, img1)
            gerar_grafico_dispersao_mediana(homog, img2)

            logger.info(f"Enviando para relatório (valores originais): {df_filtrado['VALOR TOTAL'].tolist()}")
            logger.info(f"Homogeneizados válidos: {homog}")

            finalidade_bruta = f.get("finalidade_lido", "").lower()
            if "desapropria" in finalidade_bruta:
                finalidade_tipo = "desapropriacao"
            elif "servid" in finalidade_bruta:
                finalidade_tipo = "servidao"
            else:
                finalidade_tipo = "mercado"
            if acao == "avaliar":
                from executaveis_avaliacao.utils_json import salvar_entrada_corrente_json
                salvar_entrada_corrente_json(
                    dados_imovel,
                    fatores_usuario,
                    df_amostras.to_dict(orient="records"),
                    id_execucao
                )
                return redirect(url_for('visualizar_resultados', uuid=id_execucao))



            gerar_relatorio_avaliacao_com_template(
                dados_avaliando=dados_imovel,
                dataframe_amostras_inicial=df_amostras,
                dataframe_amostras_filtrado=df_filtrado,
                indices_excluidos=idx_exc,
                amostras_excluidas=amostras_exc,
                media=media,
                desvio_padrao=dp,
                menor_valor=menor,
                maior_valor=maior,
                mediana_valor=mediana,
                valores_originais_iniciais = df_filtrado.get("VALOR TOTAL", pd.Series()).tolist(),
                valores_homogeneizados_validos=homog,
                caminho_imagem_aderencia=img1,
                caminho_imagem_dispersao=img2,
                uuid_atual=id_execucao,
                finalidade_do_laudo=finalidade_tipo,
                area_parcial_afetada=area_parcial,
                fatores_do_usuario=fatores_usuario,
                caminhos_fotos_avaliando=fotos_imovel,
                caminhos_fotos_adicionais=fotos_adicionais,
                caminhos_fotos_proprietario=fotos_proprietario,
                caminhos_fotos_planta=fotos_planta,
                caminho_template=os.path.join(BASE_DIR, "templates_doc", "Template.docx"),
                nome_arquivo_word=caminho_docx
            )
            # 3. Verificar se foi realmente criado
            if os.path.exists(caminho_docx):
                logger.info(f"✅ DOCX gerado com sucesso: {caminho_docx}")
            else:
                logger.error(f"❌ Erro: o DOCX não foi gerado em {caminho_docx}")
                        
            # 5. Gerar ZIP
            nome_zip = f"relatorio_avaliacao_{id_execucao}.zip"
            caminho_zip = os.path.join(BASE_DIR, 'static', 'arquivos', nome_zip)
            with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                logger.info(f"✅ ZIP criado em: {caminho_zip}")
                for root, dirs, files in os.walk(pasta_temp):
                    for file in files:
                        zipf.write(os.path.join(root, file), arcname=file)

            logger.info("✅ Relatório gerado com sucesso!")
            resultado = "✅ Relatório gerado com sucesso!"
            zip_download = nome_zip

            # Definir o caminho relativo ao log para o HTML
            log_path_relativo = f'logs/{os.path.basename(log_path)}'


        except Exception as e:
            erro_execucao = f"❌ Erro durante o processamento: {type(e).__name__} - {e}<br><pre>{traceback.format_exc()}</pre>"
            logger.error(erro_execucao)

    return render_template("formulario_avaliacao.html",
                           resultado=resultado,
                           erro=erro_execucao,
                           zip_download=zip_download,
                           log_path=log_path_relativo if os.path.exists(log_path) else None)



# @app.route('/memoriais-azimute-p1-p2')
# def memoriais_azimute_p1_p2():
#     return render_template('em_breve.html', titulo="MEMORIAIS_AZIMUTE_P1_P2")

# @app.route('/memoriais-angulos-internos-p1-p2')
# def memoriais_angulos_internos_p1_p2():
#     return render_template('em_breve.html', titulo="MEMORIAIS_ANGULOS_INTERNOS_P1_P2")

@app.route("/visualizar_resultados/<uuid>")
def visualizar_resultados(uuid):
    import json
    caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

    if not os.path.exists(caminho_json):
        flash("Arquivo JSON de entrada não encontrado.", "danger")
        return redirect(url_for("gerar_avaliacao"))

    with open(caminho_json, "r", encoding="utf-8") as f:
        dados = json.load(f)

    amostras = dados.get("amostras", [])
    fatores = dados.get("fatores_do_usuario", {})
    dados_avaliando = dados.get("dados_avaliando", {})

    # Calcular valor unitário por amostra ativa
    valores_ativos = [
        a["valor_total"] / a["area"]
        for a in amostras if a.get("ativo") and a.get("area", 0) > 0
    ]

    media = round(sum(valores_ativos) / len(valores_ativos), 2) if valores_ativos else 0.0

    from executaveis_avaliacao.main import intervalo_confianca_bootstrap_mediana
    amplitude_ic80 = 0.0
    if len(valores_ativos) > 1:
        li, ls = intervalo_confianca_bootstrap_mediana(valores_ativos, 1000, 0.80)
        if li > 0:
            amplitude_ic80 = round(((ls - li) / ((li + ls)/2)) * 100, 1)


    return render_template(
        "visualizar_resultados.html",
        uuid=uuid,
        amostras=amostras,
        media=media,
        amplitude_ic80=amplitude_ic80,
        dados_avaliando=dados_avaliando,
        fatores=fatores
    )
@app.route("/gerar_laudo_final/<uuid>", methods=["POST"])
def gerar_laudo_final(uuid):
    import json
    caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

    if not os.path.exists(caminho_json):
        flash("Arquivo de entrada não encontrado.", "danger")
        return redirect(url_for("gerar_avaliacao"))

    # 1. Carrega JSON
    with open(caminho_json, "r", encoding="utf-8") as f:
        dados = json.load(f)

    # 2. Atualiza estado das amostras com base nos checkboxes
    for amostra in dados["amostras"]:
        campo = f"ativo_{amostra['idx']}"
        amostra["ativo"] = campo in request.form

    # 3. Salva JSON atualizado
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)

    # 4. Filtra amostras ativas
    amostras_ativas = [
        a for a in dados["amostras"]
        if a.get("ativo") and a.get("area", 0) > 0
    ]

    if not amostras_ativas:
        flash("Nenhuma amostra ativa para gerar o laudo.", "warning")
        return redirect(url_for("visualizar_resultados", uuid=uuid))

    # 5. Executa cálculo simplificado apenas para demonstração
    valores_unitarios = [a["valor_total"] / a["area"] for a in amostras_ativas]
    media = round(sum(valores_unitarios) / len(valores_unitarios), 2)
    media_formatado = f"R$ {media:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # 6. Gera um arquivo Word simples (você pode adaptar para seu modelo completo)
    from executaveis_avaliacao.main import (
        aplicar_chauvenet_e_filtrar,
        homogeneizar_amostras,
        gerar_grafico_aderencia_totais,
        gerar_grafico_dispersao_mediana,
        gerar_relatorio_avaliacao_com_template
    )

    import pandas as pd
    df_ativas = pd.DataFrame(amostras_ativas)

    # Padroniza nomes de colunas esperadas
    df_ativas.rename(columns={
        "valor_total": "VALOR TOTAL",
        "area": "AREA TOTAL",
        "distancia_centro": "DISTANCIA CENTRO"
    }, inplace=True)


    # Aplicar Chauvenet e homogeneização
    df_filtrado, idx_exc, amostras_exc, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_ativas)
    homog = homogeneizar_amostras(df_filtrado, dados["dados_avaliando"], dados["fatores_do_usuario"], "mercado")

    pasta_saida = os.path.join("static", "arquivos", f"avaliacao_{uuid}")
    os.makedirs(pasta_saida, exist_ok=True)

    img1 = os.path.join(pasta_saida, "grafico_aderencia.png")
    img2 = os.path.join(pasta_saida, "grafico_dispersao.png")
    gerar_grafico_aderencia_totais(df_filtrado, homog, img1)
    gerar_grafico_dispersao_mediana(homog, img2)

    # Chamada final
    nome_docx = f"laudo_avaliacao_{uuid}.docx"
    caminho_docx = os.path.join(pasta_saida, nome_docx)

    gerar_relatorio_avaliacao_com_template(
        dados_avaliando=dados["dados_avaliando"],
        dataframe_amostras_inicial=df_ativas,
        dataframe_amostras_filtrado=df_filtrado,
        indices_excluidos=idx_exc,
        amostras_excluidas=amostras_exc,
        media=media,
        desvio_padrao=dp,
        menor_valor=menor,
        maior_valor=maior,
        mediana_valor=mediana,
        valores_originais_iniciais=df_filtrado["VALOR TOTAL"].tolist(),
        valores_homogeneizados_validos=homog,
        caminho_imagem_aderencia=img1,
        caminho_imagem_dispersao=img2,
        uuid_atual=uuid,
        finalidade_do_laudo="mercado",  # ou adaptar
        area_parcial_afetada=dados["dados_avaliando"].get("AREA TOTAL", 0),
        fatores_do_usuario=dados["fatores_do_usuario"],
        caminhos_fotos_avaliando=[],
        caminhos_fotos_adicionais=[],
        caminhos_fotos_proprietario=[],
        caminhos_fotos_planta=[],
        caminho_template=os.path.join(BASE_DIR, "templates_doc", "Template.docx"),
        nome_arquivo_word=caminho_docx
    )
    if os.path.exists(caminho_docx):
        logger.info(f"✅ DOCX gerado com sucesso: {caminho_docx}")
    else:
        logger.error(f"❌ Erro: o DOCX não foi gerado em {caminho_docx}")


    # 1. Criar caminho do ZIP
    nome_zip = f"pacote_avaliacao_{uuid}.zip"
    caminho_zip = os.path.join(BASE_DIR, "static", "arquivos", nome_zip)


    # Copiar o JSON de entrada para a pasta de saída
    origem_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")
    destino_json = os.path.join(pasta_saida, f"{uuid}_entrada_corrente.json")
    if os.path.exists(origem_json):
        import shutil
        shutil.copyfile(origem_json, destino_json)

    # 2. Compactar todos os arquivos da pasta
    with zipfile.ZipFile(caminho_zip, 'w') as zipf:
        for root, dirs, files in os.walk(pasta_saida):
            for file in files:
                full_path = os.path.join(root, file)
                zipf.write(full_path, arcname=file)

    # 3. Retornar ZIP ao invés do DOCX
    return send_file(caminho_zip, as_attachment=True)

   
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)


@app.route('/debug_rotas')
def debug_rotas():
    from flask import Response
    rotas = [str(rule) for rule in app.url_map.iter_rules()]
    return Response('<br>'.join(rotas), mimetype='text/html')

#finalizando#


