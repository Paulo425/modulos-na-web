from flask import (
    Flask, render_template, request, redirect, url_for, session,
    send_from_directory, send_file, flash, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from subprocess import Popen, PIPE, DEVNULL, STDOUT
import os
import json
import tempfile
from pathlib import Path
import shutil
import traceback
import sys
import fitz  # PyMuPDF
import zipfile
from pdf2image import convert_from_bytes
import io
from PIL import Image, UnidentifiedImageError
import uuid
import logging
import re
import pandas as pd  # ← inclusão imediata dessa linha resolve definitivamente
import sys
import traceback
import subprocess
import pandas as pd
from executaveis_avaliacao.main import homogeneizar_amostras



# 🔧 Configuração do logger DEFINITIVA (completa e segura)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_path = os.path.join(BASE_DIR, 'flask_app.log')

logger = logging.getLogger("app_logger")
logger.setLevel(logging.DEBUG)  # 👈 ALTERE PARA DEBUG imediatamente agora!

# Limpar handlers antigos (imprescindível agora)
if logger.hasHandlers():
    logger.handlers.clear()

# Handler arquivo
file_handler = logging.FileHandler(log_path, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Handler console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)

# Formatter robusto e completo (inclui timestamp e traceback)
formatter = logging.Formatter('%(asctime)s %(levelname)s [%(name)s] %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Confirmar imediatamente inicialização correta no log
logger.info("✅ Logger Flask configurado DEFINITIVAMENTE (nível DEBUG).")



# 📁 Diretórios base e públicos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
log_dir = os.path.join(BASE_DIR, "static", "logs")
arquivos_dir = os.path.join(BASE_DIR, "static", "arquivos")

os.makedirs(CAMINHO_PUBLICO, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)
os.makedirs(arquivos_dir, exist_ok=True)

# 🚀 Inicialização do app Flask
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)
app.secret_key = 'chave_super_secreta'
app.debug = True 
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# 🔄 Imports do módulo de usuários
from usuarios_mysql import (
    salvar_usuario_mysql, buscar_usuario_mysql, aprovar_usuario_mysql,
    excluir_usuario_mysql, listar_pendentes_mysql, listar_usuarios_mysql,
    atualizar_senha_mysql
)


def _parse_coord(coord):
    import re
    try:
        if isinstance(coord, str):
            # Remove tudo exceto dígitos, vírgula, ponto, e sinal de negativo
            coord = re.sub(r"[^\d,.\-]", "", coord).replace(",", ".").strip()
        return float(coord)
    except:
        return None


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
        sentido_poligonal = 'anti_horario' if 'sentidoPoligonal' in request.form else 'horario'

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

        # try:
        #     processo = Popen(
        #         ["python", os.path.join(BASE_DIR, "executaveis_azimute_az", "main.py"),
        #          cidade, caminho_excel, caminho_dxf],
        #         stdout=PIPE, stderr=subprocess.STDOUT, text=True
        #     )

        #     log_lines = []
        #     with open(log_path, 'w', encoding='utf-8') as log_file:
        #         for linha in processo.stdout:
        #             log_file.write(linha)
        #             if len(log_lines) < 500:
        #                 log_lines.append(linha)
        #             print("🖨️", linha.strip())

        #     processo.wait()

        #     if processo.returncode == 0:
        #         resultado = "✅ Processamento concluído com sucesso!"
        #     else:
        #         erro_execucao = f"❌ Erro na execução:<br><pre>{''.join(log_lines)}</pre>"

        # except Exception as e:
        #     erro_execucao = f"❌ Erro inesperado:<br><pre>{type(e).__name__}: {str(e)}</pre>"

        try:
            comando = [
                sys.executable,
                os.path.join(BASE_DIR, "executaveis_azimute_p1_p2", "main.py"),
                cidade, caminho_excel, caminho_dxf, sentido_poligonal
            ]

            logger.info(f"Comando enviado ao subprocess: {comando}")

            processo = Popen(
                comando,
                stdout=PIPE, stderr=STDOUT, text=True
            )

            try:
                saida, _ = processo.communicate(timeout=300)
                logger.info(f"Saída do subprocess:\n{saida}")
            except TimeoutExpired:
                processo.kill()
                saida, _ = processo.communicate()
                logger.error(f"Subprocess atingiu timeout. Saída parcial:\n{saida}")

        except Exception as e:
            logger.error(f"Erro fatal ao executar subprocess: {e}")


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
        sentido_poligonal = 'anti_horario' if 'sentidoPoligonal' in request.form else 'horario'
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

       

        # try:
        #     processo = Popen(
        #         [sys.executable, os.path.join(BASE_DIR, "executaveis_angulo_p1_p2", "main.py"),
        #         cidade, caminho_excel, caminho_dxf],
        #         stdout=PIPE, stderr=STDOUT, text=True
        #     )

        #     log_lines = []
        #     with open(log_path, 'w', encoding='utf-8') as log_file:
        #         for linha in processo.stdout:
        #             log_file.write(linha)
        #             if len(log_lines) < 500:
        #                 log_lines.append(linha)
        #             print("🖨️", linha.strip())

        #     processo.wait()

        #     if processo.returncode == 0:
        #         resultado = "✅ Processamento concluído com sucesso!"
        #     else:
        #         erro_execucao = f"❌ Erro na execução:<br><pre>{''.join(log_lines)}</pre>"

        # except Exception as e:
        #     erro_execucao = f"❌ Erro inesperado:<br><pre>{type(e).__name__}: {str(e)}</pre>"


        try:
            comando = [
                sys.executable,
                os.path.join(BASE_DIR, "executaveis_angulo_p1_p2", "main.py"),
                cidade, caminho_excel, caminho_dxf, sentido_poligonal
            ]

            logger.info(f"Comando enviado ao subprocess: {comando}")

            processo = Popen(
                comando,
                stdout=PIPE, stderr=STDOUT, text=True
            )

            try:
                saida, _ = processo.communicate(timeout=300)
                logger.info(f"Saída do subprocess:\n{saida}")
            except TimeoutExpired:
                processo.kill()
                saida, _ = processo.communicate()
                logger.error(f"Subprocess atingiu timeout. Saída parcial:\n{saida}")

        except Exception as e:
            logger.error(f"Erro fatal ao executar subprocess: {e}")




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
from subprocess import Popen, PIPE, STDOUT, CalledProcessError, TimeoutExpired
@app.route('/memorial_azimute_p1_p2', methods=['GET', 'POST'])
def gerar_memorial_azimute_p1_p2():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultado = erro_execucao = log_relativo = None
    amostras_homog = []

    zip_download = None

    if request.method == 'POST':
        cidade = request.form['cidade'].strip()
        sentido_poligonal = 'anti_horario' if 'sentidoPoligonal' in request.form else 'horario'
        logger.info(f"Valor recebido do checkbox (sentidoPoligonal): {request.form.get('sentidoPoligonal')}")
        logger.info(f"Sentido poligonal interpretado no Flask: {sentido_poligonal}")

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
            comando = [
                sys.executable,
                os.path.join(BASE_DIR, "executaveis_azimute_p1_p2", "main.py"),
                cidade, caminho_excel, caminho_dxf, sentido_poligonal
            ]

            logger.info(f"Comando enviado ao subprocess: {comando}")

            processo = Popen(
                comando,
                stdout=PIPE, stderr=STDOUT, text=True
            )

            try:
                saida, _ = processo.communicate(timeout=300)
                logger.info(f"Saída do subprocess:\n{saida}")
            except TimeoutExpired:
                processo.kill()
                saida, _ = processo.communicate()
                logger.error(f"Subprocess atingiu timeout. Saída parcial:\n{saida}")

        except Exception as e:
            logger.error(f"Erro fatal ao executar subprocess: {e}")

            
            
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

    amostras_homog = []
    
    #logger = logging.getlogger(__name__)  # ← ajuste definitivo aqui!
    logger.debug("🚀 Iniciando rota gerar_avaliacao()")


    try:
        logger.debug("Iniciando rota gerar_avaliacao()")

        if 'usuario' not in session:
            return redirect(url_for('login'))

        resultado = erro_execucao = zip_download = log_relativo = None

        if request.method == "POST":
            logger.info("🔧 Início da execução do bloco POST em /avaliacoes")

            # Indispensável! Identifica o botão clicado pelo usuário
            acao = request.form.get("acao", "").lower()
            logger.debug(f"Ação recebida: {acao}")

            # Indispensável! Verifica o envio da planilha Excel
            if "planilha_excel" not in request.files:
                logger.error("❌ ERRO: O arquivo 'planilha_excel' não foi enviado!")
                return "Erro: arquivo planilha_excel faltando!", 400

            excel_file = request.files["planilha_excel"]
            if excel_file.filename == '':
                logger.error("❌ ERRO: Arquivo planilha_excel vazio ou nome inválido.")
                return "Erro: arquivo planilha_excel vazio.", 400

            try:
                from werkzeug.utils import secure_filename
                import uuid, zipfile

                # 1. Criação de diretório temporário para essa execução
                id_execucao = str(uuid.uuid4())[:8]
                pasta_execucao = f'avaliacao_{id_execucao}'
                pasta_temp = os.path.join(BASE_DIR, 'static', 'arquivos', pasta_execucao)
                os.makedirs(pasta_temp, exist_ok=True)

                # 2. Salvar arquivo Excel recebido
                caminho_planilha = os.path.join(pasta_temp, "planilha.xlsx")
                excel_file.save(caminho_planilha)
                logger.info(f"✅ Planilha salva: {caminho_planilha}")

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

                # NOVA LINHA: Crie uma nova chave sem alterar AREA TOTAL original
               
                #FAZ O TRATAMENTO EM TODAS AS COORDENADAS DO EXCEL********************FOI AQUI RETIRADO TESTE TEMPORATIRO*
                df_amostras, dados_imovel = ler_planilha_excel(caminho_planilha)

                # Adicione imediatamente após essa linha:
                df_amostras["idx"] = df_amostras["AM"].astype(int)

                # ▼▼▼ CÁLCULO DO VALOR UNITÁRIO MÉDIO ▼▼▼
                # Supondo que você ainda não rodou homogeneização, então calcule manualmente:
                valores_unitarios = [
                    row["VALOR TOTAL"] / row["AREA TOTAL"] if row["AREA TOTAL"] > 0 else 0
                    for _, row in df_amostras.iterrows()
                ]
                valores_unitarios = [v for v in valores_unitarios if v > 0]  # filtra apenas valores realmente válidos

                valor_unitario_medio = sum(valores_unitarios) / len(valores_unitarios) if valores_unitarios else 0
                dados_imovel["valor_unitario_medio"] = valor_unitario_medio

                # ▲▲▲ FIM DO BLOCO DE CÁLCULO ▲▲▲

                # NOVA LINHA: Pegue a área digitada pelo usuário no input

                area_parcial_afetada = float(request.form.get("area_parcial_afetada", "0").replace(".", "").replace(",", "."))
                dados_imovel["AREA_PARCIAL_AFETADA"] = float(area_parcial_afetada)


                # Função que remove graus e espaços
               # Limpeza e conversão robusta das coordenadas do imóvel avaliado
                dados_imovel["LATITUDE"] = _parse_coord(dados_imovel.get("LATITUDE"))
                dados_imovel["LONGITUDE"] = _parse_coord(dados_imovel.get("LONGITUDE"))

                # Limpeza robusta das coordenadas das amostras
                for col in ["LATITUDE", "LONGITUDE"]:
                    if col in df_amostras.columns:
                        df_amostras[col] = df_amostras[col].apply(_parse_coord)
                # Logs detalhados
                logger.info(f"Coordenadas limpas imóvel: LATITUDE={dados_imovel['LATITUDE']}, LONGITUDE={dados_imovel['LONGITUDE']}")
                logger.info(f"Primeiras linhas df_amostras após limpeza:\n{df_amostras[['LATITUDE', 'LONGITUDE']].head()}")

                #**********************************************************************
                logger.info(f"df_amostras.head():\n{df_amostras.head()}")
                logger.info(f"dados_imovel: {dados_imovel}")
                # AQUI RETIRADO TEMPORARIAMENTE
                df_filtrado, idx_exc, amostras_exc, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_amostras)
                

                logger.info(f"df_filtrado.head():\n{df_filtrado.head()}")
                logger.info(f"Média: {media}, Mediana: {mediana}")
                #AQUI RETIRADO TEMPORARIAMENTE
                #homog = homogeneizar_amostras(df_filtrado, dados_imovel, fatores_usuario, "mercado")
                amostras_homog = homogeneizar_amostras(df_filtrado, dados_imovel, fatores_usuario, "mercado")
                

                # Separando listas após homogeneização (novo)
                lista_valores_unitarios = [a["valor_unitario"] for a in amostras_homog]
                lista_residuos_relativos = [a["residuo_rel"] for a in amostras_homog]
                lista_residuos_dp = [a["residuo_dp"] for a in amostras_homog]
                img1 = os.path.join(pasta_temp, "grafico_aderencia.png")
                img2 = os.path.join(pasta_temp, "grafico_dispersao.png")
               

                gerar_grafico_aderencia_totais(df_filtrado, [a["valor_unitario"] for a in amostras_homog], img1)

                # solução imediata e recomendada para gerar_avaliacao
                idx_todas_amostras = df_amostras["idx"].tolist()
                gerar_grafico_dispersao_mediana(
                    df_filtrado,
                    [a["valor_unitario"] for a in amostras_homog],
                    img2,
                    idx_todas_amostras,  # amostras iniciais
                    [],                  # nenhuma retirada manual
                    []                   # nenhuma retirada Chauvenet
                )

                logger.info(f"Enviando para relatório (valores originais): {df_filtrado['VALOR TOTAL'].tolist()}")
                logger.info(f"Homogeneizados válidos: {amostras_homog}")

                finalidade_bruta = f.get("finalidade_lido", "").lower()
                if "desapropria" in finalidade_bruta:
                    finalidade_tipo = "desapropriacao"
                elif "servid" in finalidade_bruta:
                    finalidade_tipo = "servidao"
                else:
                    finalidade_tipo = "mercado"
                if acao == "avaliar":
                    from executaveis_avaliacao.utils_json import salvar_entrada_corrente_json
                    
                    lista_amostras = []
                    for _, linha in df_amostras.iterrows():
                        area = float(linha.get("AREA TOTAL", 0))
                        valor_total = float(linha.get("VALOR TOTAL", 0))

                        latitude = linha.get("LATITUDE")
                        longitude = linha.get("LONGITUDE")

                        logger.info(f"Latitude final: {latitude}, Longitude final: {longitude}")

                        lista_amostras.append({
                            "idx": linha.get("AM", ""),
                            "valor_total": valor_total,
                            "area": area,
                            "LATITUDE": latitude,
                            "LONGITUDE": longitude,
                            "cidade": linha.get("CIDADE", ""),
                            "fonte": linha.get("FONTE", ""),
                            "ativo": True
                        })


                    
                    salvar_entrada_corrente_json(
                        uuid_execucao=id_execucao,
                        dados_avaliando=dados_imovel,
                        fatores_do_usuario=fatores_usuario,
                        amostras=lista_amostras,
                        fotos_imovel=fotos_imovel,
                        fotos_adicionais=fotos_adicionais,
                        fotos_proprietario=fotos_proprietario,
                        fotos_planta=fotos_planta,
                        base_dir=BASE_DIR,  # opcional, mas bom para garantir o caminho correto
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
                    valores_homogeneizados_validos=lista_valores_unitarios,
                    caminho_imagem_aderencia=img1,
                    caminho_imagem_dispersao=img2,
                    uuid_atual=id_execucao,
                    finalidade_do_laudo=finalidade_tipo,
                    area_parcial_afetada=area_parcial_afetada, # aqui vai o correto
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
                # NOVO BLOCO CRÍTICO: logar em arquivo adicional
                from pathlib import Path

                LOG_DIR = Path(__file__).parent / "static" / "logs"
                LOG_DIR.mkdir(parents=True, exist_ok=True)

                with open(LOG_DIR / "erro_critico.log", "a") as f:
                    f.write(erro_execucao + "\n")


        return render_template(
            "formulario_avaliacao.html",
            resultado=resultado,
            erro=erro_execucao,
            zip_download=zip_download,
            log_path=log_path_relativo if 'log_path_relativo' in locals() and log_path_relativo and os.path.exists(log_path_relativo) else None,
            amostras=amostras_homog if request.method == "POST" else []  # ← Correção definitiva aqui
        )


    except Exception as e:
        logger.exception(f"🚨 Erro ao iniciar processamento: {e}")
        return f"Erro interno ao iniciar processamento: {str(e)}", 500


# @app.route("/visualizar_resultados/<uuid>")
# def visualizar_resultados(uuid):

#     if request.method == "POST" and request.form.get("acao") == "gerar_laudo":
#         return gerar_laudo_final(uuid)

#     caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

#     logger.info(f"✅ Iniciando visualizar_resultados() para UUID: {uuid}")
#     logger.info(f"📂 Caminho JSON: {caminho_json}")

#     if not os.path.exists(caminho_json):
#         logger.error("❌ Arquivo JSON não encontrado.")
#         flash("Arquivo JSON de entrada não encontrado.", "danger")
#         return redirect(url_for("gerar_avaliacao"))

#     try:
#         with open(caminho_json, "r", encoding="utf-8") as f:
#             dados = json.load(f)
#         logger.info("📌 JSON carregado com sucesso.")

#         amostras = dados.get("amostras", [])
#         fatores = dados.get("fatores_do_usuario", {})
#         dados_avaliando = dados.get("dados_avaliando", {})

#         # Passo elegante: só passe amostras prontas para o template!
        

#         amostras_ativas = [a for a in amostras if a.get("ativo") and a.get("area", 0) > 0]
#         df_ativas = pd.DataFrame(amostras_ativas)
#         df_ativas.rename(columns={
#             "valor_total": "VALOR TOTAL",
#             "area": "AREA TOTAL",
#             "distancia_centro": "DISTANCIA CENTRO"
#         }, inplace=True)

#         # >>> INSIRA ESTE BLOCO AQUI <<<
#         valores_unitarios = [
#             row["VALOR TOTAL"] / row["AREA TOTAL"] if row["AREA TOTAL"] > 0 else 0
#             for _, row in df_ativas.iterrows()
#         ]
#         valor_unitario_medio = sum(valores_unitarios) / len([v for v in valores_unitarios if v > 0]) if valores_unitarios else 0
#         dados_avaliando["valor_unitario_medio"] = valor_unitario_medio


#         amostras_prontas = homogeneizar_amostras(
#             df_ativas, 
#             dados_avaliando, 
#             fatores, 
#             finalidade_do_laudo="mercado"
#         )
       
#         # Calcula os valores ativos e a média
#         valores_ativos = [a["valor_unitario"] for a in amostras_prontas if a.get("area", 0) > 0]
#         if valores_ativos:
#             media = round(sum(valores_ativos) / len(valores_ativos), 2)
#             logger.info(f"📊 Média calculada: {media}")
#         else:
#             media = 0.0
#             logger.warning("⚠️ Nenhum valor ativo encontrado para média.")

#         from executaveis_avaliacao.main import intervalo_confianca_bootstrap_mediana
#         amplitude_ic80 = 0.0
#         if len(valores_ativos) > 1:
#             logger.info("📌 Iniciando cálculo do intervalo de confiança bootstrap.")
#             li, ls = intervalo_confianca_bootstrap_mediana(valores_ativos, 1000, 0.80)
#             logger.info(f"📌 IC 80% calculado: LI={li}, LS={ls}")
#             if li > 0:
#                 amplitude_ic80 = round(((ls - li) / ((li + ls)/2)) * 100, 1)
#                 logger.info(f"📊 Amplitude IC 80%: {amplitude_ic80}%")
#             else:
#                 logger.warning("⚠️ LI do intervalo é menor ou igual a zero.")
#         else:
#             logger.warning("⚠️ Não há valores suficientes para calcular IC 80%.")

#     except Exception as erro:
#         logger.exception(f"🚨 Exceção capturada em visualizar_resultados: {erro}")

        
#         erro_completo = traceback.format_exc()
#         erro_arquivo = os.path.join(BASE_DIR, "erro_avaliacao.txt")
#         with open(erro_arquivo, "w", encoding="utf-8") as arquivo_erro:
#             arquivo_erro.write(erro_completo)
#         flash(f"Erro detalhado capturado: {erro}", "danger")
#         return redirect(url_for("gerar_avaliacao"))

#     logger.info("🚩 Renderizando template visualizar_resultados.html")

#     # Sempre passar amostras homogeneizadas para o template — garante todos os campos derivados e evita erros de atributo.
#     return render_template(
#         "visualizar_resultados.html",
#         uuid=uuid,
#         amostras=amostras_prontas,
#         media=media,
#         amplitude_ic80=amplitude_ic80,
#         dados_avaliando=dados_avaliando,
#         fatores=fatores
#     )

# @app.route("/visualizar_resultados/<uuid>", methods=["GET", "POST"])
# def visualizar_resultados(uuid):
#     # 🔁 Se veio submit do botão "Gerar Laudo Final", delega direto
#     if request.method == "POST" and request.form.get("acao") == "gerar_laudo":
#         return gerar_laudo_final(uuid)

#     caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

#     logger.info(f"✅ Iniciando visualizar_resultados() para UUID: {uuid}")
#     logger.info(f"📂 Caminho JSON: {caminho_json}")

#     if not os.path.exists(caminho_json):
#         logger.error("❌ Arquivo JSON não encontrado.")
#         flash("Arquivo JSON de entrada não encontrado.", "danger")
#         return redirect(url_for("gerar_avaliacao"))

#     try:
#         with open(caminho_json, "r", encoding="utf-8") as f:
#         dados = json.load(f)

#         amostras = dados.get("amostras", []) or []
#         fatores = dados.get("fatores_do_usuario", {}) or {}
#         dados_avaliando = dados.get("dados_avaliando", {}) or {}

#         # Só ativas e com área
#         amostras_ativas = [a for a in amostras if a.get("ativo") and (a.get("area") or 0) > 0]

#         # OPTIM: se não há ativas, evita custos e devolve rápido
#         if not amostras_ativas:
#             logger.warning("⚠️ Nenhuma amostra ativa encontrada para a tela iterativa.")
#             return render_template(
#                 "visualizar_resultados.html",
#                 uuid=uuid,
#                 amostras=[],
#                 media=0.0,
#                 amplitude_ic80=0.0,
#                 dados_avaliando=dados_avaliando,
#                 fatores=fatores,
#             )

#         # OPTIM: monta apenas as colunas necessárias sem DataFrame pesado
#         # ---------- extrai campos ----------
#         valores_totais = [float(a.get("valor_total", 0) or 0) for a in amostras_ativas]
#         areas          = [float(a.get("area", 0) or 0)         for a in amostras_ativas]
#         idxs           = [int(a.get("idx", 0) or 0)            for a in amostras_ativas]
#         dist_centro    = [float(a.get("distancia_centro", 0) or 0) for a in amostras_ativas]

#         # >>> NOVO: trazer PAV e ACESS das amostras (se vierem None, o cálculo usará o do avaliado)
#         pav_list   = [get_multi(a, "PAVIMENTACAO?", "PAVIMENTAÇÃO?", "PAVIMENTACAO ?") for a in amostras_ativas]
#         acess_list = [get_multi(a, "ACESSIBILIDADE?", "ACESSIBILIDADE ?")                for a in amostras_ativas]

#         # Valor unitário médio (como estava)
#         vu_list = [(vt / ar) if ar > 0 else 0.0 for vt, ar in zip(valores_totais, areas)]
#         vu_validos = [v for v in vu_list if v > 0]
#         dados_avaliando["valor_unitario_medio"] = (sum(vu_validos) / len(vu_validos)) if vu_validos else 0.0

#         # Reconstrói um "mini-DF" só com as colunas esperadas pela homogeneização
#         # (Se sua função aceita dict/lista, melhor ainda; se exige DataFrame, montamos leve)
#         import pandas as pd
#         df_ativas = pd.DataFrame({
#             "VALOR TOTAL": valores_totais,
#             "AREA TOTAL": areas,
#             "DISTANCIA CENTRO": dist_centro,
#             "idx": idxs,
#             # >>> NOVO: colunas de fatores que a homogeneização precisa ler
#             "PAVIMENTACAO?": pav_list,
#             "ACESSIBILIDADE?": acess_list,
#         })
#         # (opcional) logs para verificar o que chegou
#         logger.info(f"RAW PAV (primeiras 5): {pav_list[:5]}")
#         logger.info(f"RAW ACESS (primeiras 5): {acess_list[:5]}")
#         logger.info(f"Avaliando PAV='{dados_avaliando.get('PAVIMENTACAO?')}', ACESS='{dados_avaliando.get('ACESSIBILIDADE?')}'")
#         # OPTIM: chama homogeneizar_amostras apenas se houver linhas
#         if len(df_ativas) == 0:
#             amostras_prontas = []
#         else:
#             amostras_prontas = homogeneizar_amostras(
#                 df_ativas,
#                 dados_avaliando,
#                 fatores,
#                 finalidade_do_laudo="mercado",
#             )

#         # Média dos valores ativos (pós-homog.)
#         valores_ativos = [float(a.get("valor_unitario", 0) or 0) for a in amostras_prontas if (a.get("area", 0) or 0) > 0]
#         if valores_ativos:
#             media = round(sum(valores_ativos) / len(valores_ativos), 2)
#         else:
#             media = 0.0

#         # OPTIM: Bootstrap é caro; use 400 em vez de 1000
#         amplitude_ic80 = 0.0
#         if len(valores_ativos) > 1:
#             from executaveis_avaliacao.main import intervalo_confianca_bootstrap_mediana
#             li, ls = intervalo_confianca_bootstrap_mediana(valores_ativos, 400, 0.80)
#             if li > 0:
#                 amplitude_ic80 = round(((ls - li) / ((li + ls) / 2)) * 100, 1)

#     except Exception as erro:
#         logger.exception(f"🚨 Exceção capturada em visualizar_resultados: {erro}")
#         import traceback
#         erro_completo = traceback.format_exc()
#         erro_arquivo = os.path.join(BASE_DIR, "erro_avaliacao.txt")
#         with open(erro_arquivo, "w", encoding="utf-8") as arquivo_erro:
#             arquivo_erro.write(erro_completo)
#         flash(f"Erro detalhado capturado: {erro}", "danger")
#         return redirect(url_for("gerar_avaliacao"))

#     logger.info("🚩 Renderizando template visualizar_resultados.html")

#     return render_template(
#         "visualizar_resultados.html",
#         uuid=uuid,
#         amostras=amostras_prontas,
#         media=media,
#         amplitude_ic80=amplitude_ic80,
#         dados_avaliando=dados_avaliando,
#         fatores=fatores,
#     )

@app.route("/visualizar_resultados/<uuid>", methods=["GET", "POST"])
def visualizar_resultados(uuid):
    """
    Tela intermediária de controle.
    Lê TUDO do snapshot JSON (schema v2) e NUNCA de variáveis em memória.
    """
    if request.method == "POST" and request.form.get("acao") == "gerar_laudo":
        return gerar_laudo_final(uuid)

    try:
        from executaveis_avaliacao.utils_json import carregar_entrada_corrente_json
        from executaveis_avaliacao.main import (
            homogeneizar_amostras,
            intervalo_confianca_bootstrap_mediana,
        )
        import pandas as pd
        import numpy as np
        import random, os

        logger.info(f"✅ Iniciando visualizar_resultados() para UUID: {uuid}")

        # Carrega snapshot
        data = carregar_entrada_corrente_json(uuid)

        amostras         = data.get("amostras", []) or []
        fatores          = data.get("fatores_do_usuario", {}) or {}
        dados_avaliando  = data.get("dados_avaliando", {}) or {}
        cfg              = data.get("config_modelo", {}) or {}
        params           = cfg.get("parametros", {}) or {}

        # Seed
        seed = cfg.get("random_seed")
        if isinstance(seed, int):
            random.seed(seed)
            np.random.seed(seed)

        # Filtra ativas
        amostras_ativas = [
            a for a in amostras
            if bool(a.get("ativo", True)) and float(a.get("area", 0) or 0) > 0
        ]
        if not amostras_ativas:
            logger.warning("⚠️ Nenhuma amostra ativa encontrada para a tela iterativa.")
            return render_template(
                "visualizar_resultados.html",
                uuid=uuid,
                amostras=[],
                media=0.0,
                amplitude_ic80=0.0,
                dados_avaliando=dados_avaliando,
                fatores=fatores,
            )

        # Cria DF direto das amostras
        df_ativas = pd.DataFrame(amostras_ativas)
        df_ativas.rename(columns={
            "valor_total": "VALOR TOTAL",
            "area": "AREA TOTAL",
        }, inplace=True)
        if "idx" not in df_ativas.columns:
            df_ativas["idx"] = 0
        if "DISTANCIA CENTRO" not in df_ativas.columns:
            df_ativas["DISTANCIA CENTRO"] = 0.0

        # Valor unitário médio bruto
        vu_list    = [(vt / ar) if ar > 0 else 0.0 for vt, ar in zip(df_ativas["VALOR TOTAL"], df_ativas["AREA TOTAL"])]
        vu_validos = [v for v in vu_list if v > 0]
        dados_avaliando["valor_unitario_medio"] = (sum(vu_validos) / len(vu_validos)) if vu_validos else 0.0

        # Homogeneização
        amostras_prontas = homogeneizar_amostras(
            df_ativas,
            dados_avaliando,
            fatores,
            finalidade_do_laudo="mercado",
        )

        # Média pós-homog.
        valores_ativos = [
            float(a.get("valor_unitario", 0) or 0)
            for a in amostras_prontas
            if float(a.get("area", 0) or 0) > 0
        ]
        media = round(sum(valores_ativos) / len(valores_ativos), 2) if valores_ativos else 0.0

        # IC 80%
        bootstrap_n = int(params.get("bootstrap_n", 400) or 400)
        amplitude_ic80 = 0.0
        if len(valores_ativos) > 1:
            li, ls = intervalo_confianca_bootstrap_mediana(valores_ativos, bootstrap_n, 0.80)
            if li > 0:
                amplitude_ic80 = round(((ls - li) / ((li + ls) / 2)) * 100, 1)

        logger.info("🚩 Renderizando template visualizar_resultados.html")
        return render_template(
            "visualizar_resultados.html",
            uuid=uuid,
            amostras=amostras_prontas,
            media=media,
            amplitude_ic80=amplitude_ic80,
            dados_avaliando=dados_avaliando,
            fatores=fatores,
        )

    except FileNotFoundError as e:
        logger.error(f"❌ Snapshot JSON não encontrado para UUID={uuid}: {e}")
        flash("Arquivo JSON de entrada não encontrado. Refaça a etapa de entrada.", "danger")
        return redirect(url_for("gerar_avaliacao"))
    except Exception as erro:
        logger.exception(f"🚨 Exceção capturada em visualizar_resultados: {erro}")
        flash(f"Erro detalhado capturado: {erro}", "danger")
        return redirect(url_for("gerar_avaliacao"))



@app.route("/gerar_laudo_final/<uuid>", methods=["POST"])
def gerar_laudo_final(uuid):
    if request.form.get("acao") != "gerar_laudo":
        flash("Ação inválida ou acesso direto sem clique autorizado.", "warning")
        return redirect(url_for("visualizar_resultados", uuid=uuid))

    global logger
    caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

    if not os.path.exists(caminho_json):
        flash("Arquivo de entrada não encontrado.", "danger")
        return redirect(url_for("gerar_avaliacao"))

    # Carrega JSON
    with open(caminho_json, "r", encoding="utf-8") as f:
        dados = json.load(f)
        fotos_imovel = dados.get("fotos_imovel", [])
        fotos_adicionais = dados.get("fotos_adicionais", [])
        fotos_proprietario = dados.get("fotos_proprietario", [])
        fotos_planta = dados.get("fotos_planta", [])
        area_parcial_afetada = float(dados["dados_avaliando"].get("AREA_PARCIAL_AFETADA", 0))

    
    # Atualiza estado das amostras (apenas se o front realmente enviar as chaves)
    chaves_postadas = {k for k in request.form.keys() if k.startswith("ativo_")}
    logger.info(f"🔎 Campos recebidos no POST: {sorted(request.form.keys())}")

    if chaves_postadas:
        # Só sobrescreve se o POST trouxe os "ativo_*"
        for amostra in dados["amostras"]:
            campo = f"ativo_{amostra['idx']}"
            amostra["ativo"] = campo in request.form
            # Obs.: checkbox não enviado = False
    else:
        logger.warning("⚠️ Nenhuma chave 'ativo_*' recebida. Mantendo estados 'ativo' do JSON.")


    # Salva JSON atualizado
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)

    # Importações obrigatórias ANTES de usar as funções
    import pandas as pd
    from executaveis_avaliacao.main import (
        aplicar_chauvenet_e_filtrar,
        homogeneizar_amostras,
        gerar_grafico_aderencia_totais,
        gerar_grafico_dispersao_mediana,
        gerar_relatorio_avaliacao_com_template
    )

    # Continuar sem alterações:
    ativos_frontend = [a["idx"] for a in dados["amostras"] if a.get("ativo", False)]
    amostras_usuario_retirou = [a["idx"] for a in dados["amostras"] if not a.get("ativo", False)]

    # Filtra amostras ativas
    amostras_ativas = [a for a in dados["amostras"] if a.get("ativo") and a.get("area", 0) > 0]

    if not amostras_ativas:
        flash("Nenhuma amostra ativa para gerar o laudo.", "warning")
        return redirect(url_for("visualizar_resultados", uuid=uuid))

    df_ativas = pd.DataFrame(amostras_ativas)
    df_ativas.rename(columns={
        "valor_total": "VALOR TOTAL",
        "area": "AREA TOTAL",
        "distancia_centro": "DISTANCIA CENTRO"
    }, inplace=True)


    # ▼▼▼ Calcule o valor_unitario_medio e adicione ao dicionário ▼▼▼
    valores_unitarios = [
        row["VALOR TOTAL"] / row["AREA TOTAL"] if row["AREA TOTAL"] > 0 else 0
        for _, row in df_ativas.iterrows()
    ]
    valor_unitario_medio = sum(valores_unitarios) / len([v for v in valores_unitarios if v > 0]) if valores_unitarios else 0
    dados["dados_avaliando"]["valor_unitario_medio"] = valor_unitario_medio
    # ▲▲▲ FIM DO BLOCO ▲▲▲

    # Aplicar Chauvenet e homogeneização
    df_filtrado, idx_exc, amostras_exc, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_ativas)
    amostras_homog = homogeneizar_amostras(df_filtrado, dados["dados_avaliando"], dados["fatores_do_usuario"], "mercado")
    

    # Alinhar tudo pelos índices reais (idx), não por posições do enumerate
    idx_filtrados = df_filtrado["idx"].astype(int).tolist()
    ativos_frontend_set = set(int(i) for i in ativos_frontend)

    amostras_chauvenet_retirou = [i for i in ativos_frontend_set if i not in idx_filtrados]

    # Mapa idx → valor_unitario vindo da homogeneização
    map_vu_por_idx = {int(a["idx"]): float(a["valor_unitario"]) for a in amostras_homog}

    # Y do scatter deve ter MESMA ordem/quantidade que o X que a função calcula (df_filtrado ∩ ativos_frontend)
    indices_ativos_alinhados = [i for i in idx_filtrados if i in ativos_frontend_set and i in map_vu_por_idx]
    valores_unit_ativos = [map_vu_por_idx[i] for i in indices_ativos_alinhados]

    pasta_saida = os.path.join("static", "arquivos", f"avaliacao_{uuid}")

    os.makedirs(pasta_saida, exist_ok=True)

    img1 = os.path.join(pasta_saida, "grafico_aderencia_iterativo.png")
    img2 = os.path.join(pasta_saida, "grafico_dispersao_iterativo.png")


    gerar_grafico_dispersao_mediana(
        df_filtrado,
        valores_unit_ativos,
        img2,
        ativos_frontend,
        amostras_usuario_retirou,
        amostras_chauvenet_retirou
    )

    finalidade_digitada = dados["fatores_do_usuario"].get("finalidade_descricao", "").strip().lower()

    if "desapropria" in finalidade_digitada:
        finalidade_do_laudo = "desapropriacao"
    elif "servid" in finalidade_digitada:
        finalidade_do_laudo = "servidao"
    else:
        finalidade_do_laudo = "mercado"

    caminho_docx = os.path.join(pasta_saida, f"laudo_avaliacao_{uuid}.docx")

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
        valores_homogeneizados_validos=amostras_homog,
        caminho_imagem_aderencia=img1,
        caminho_imagem_dispersao=img2,
        uuid_atual=uuid,
        finalidade_do_laudo=finalidade_do_laudo,
        area_parcial_afetada=area_parcial_afetada,
        fatores_do_usuario=dados["fatores_do_usuario"],
        caminhos_fotos_avaliando=fotos_imovel,
        caminhos_fotos_adicionais=fotos_adicionais,
        caminhos_fotos_proprietario=fotos_proprietario,
        caminhos_fotos_planta=fotos_planta,
        caminho_template=os.path.join(BASE_DIR, "templates_doc", "Template.docx"),
        nome_arquivo_word=caminho_docx
    )

    if os.path.exists(caminho_docx):
        logger.info(f"✅ DOCX gerado com sucesso: {caminho_docx}")
    else:
        logger.error(f"❌ Erro: o DOCX não foi gerado em {caminho_docx}")

    nome_zip = f"pacote_avaliacao_{uuid}.zip"
    caminho_zip = os.path.join(BASE_DIR, "static", "arquivos", nome_zip)

    origem_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")
    destino_json = os.path.join(pasta_saida, f"{uuid}_entrada_corrente.json")

    if os.path.exists(origem_json):
        import shutil
        shutil.copyfile(origem_json, destino_json)

    with zipfile.ZipFile(caminho_zip, 'w') as zipf:
        for root, dirs, files in os.walk(pasta_saida):
            for file in files:
                full_path = os.path.join(root, file)
                zipf.write(full_path, arcname=file)

    return send_file(caminho_zip, as_attachment=True)



# @app.route("/calcular_valores_iterativos/<uuid>", methods=["POST"])
# def calcular_valores_iterativos(uuid):
#     import json, os
#     import numpy as np
#     import pandas as pd
#     from flask import jsonify, request, url_for
#     from executaveis_avaliacao.main import (
#         aplicar_chauvenet_e_filtrar,
#         homogeneizar_amostras,
#         intervalo_confianca_bootstrap_mediana,
#         gerar_grafico_dispersao_mediana,
#         gerar_grafico_aderencia_totais,
#     )

#     try:
#         logger.info("🚀 Rota calcular_valores_iterativos iniciada")

#         caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")

#         if not os.path.exists(caminho_json):
#             logger.error(f"❌ Arquivo não encontrado: {caminho_json}")
#             return jsonify({"erro": "Arquivo de entrada não encontrado."}), 400

#         with open(caminho_json, "r", encoding="utf-8") as f:
#             dados = json.load(f)

#         ativos_frontend = request.json.get("ativos", [])
#         ativos_frontend = [int(idx) for idx in ativos_frontend]

#         amostras_usuario_retirou = [
#             int(a["idx"]) for a in dados["amostras"] if int(a["idx"]) not in ativos_frontend
#         ]

#         df_ativas = pd.DataFrame([a for a in dados["amostras"] if int(a["idx"]) in ativos_frontend])
#         df_ativas.rename(columns={"valor_total": "VALOR TOTAL", "area": "AREA TOTAL"}, inplace=True)
       
#         # === NOVO: garantir PAVIMENTACAO? e ACESSIBILIDADE? nas amostras ===
#         def get_multi(d, *keys):
#             for k in keys:
#                 if k in d and d[k] not in (None, "", "NaN"):
#                     return d[k]
#             return None

#         pav_list = [get_multi(a, "PAVIMENTACAO?", "PAVIMENTAÇÃO?", "PAVIMENTACAO ?") for a in dados["amostras"] if int(a["idx"]) in ativos_frontend]
#         acess_list = [get_multi(a, "ACESSIBILIDADE?", "ACESSIBILIDADE ?") for a in dados["amostras"] if int(a["idx"]) in ativos_frontend]

#         # Se não existir na lista (porque o JSON não traz), entra None; a lógica de homogenização já trata:
#         df_ativas["PAVIMENTACAO?"]  = pav_list
#         df_ativas["ACESSIBILIDADE?"] = acess_list

#         # (opcional) logs de sanidade
#         logger.info(f"PAV (amostras ativas) head: {list(df_ativas['PAVIMENTACAO?'][:5])}")
#         logger.info(f"ACESS (amostras ativas) head: {list(df_ativas['ACESSIBILIDADE?'][:5])}")
#         logger.info(f"Avaliando PAV='{dados['dados_avaliando'].get('PAVIMENTACAO?')}', "
#                     f"ACESS='{dados['dados_avaliando'].get('ACESSIBILIDADE?')}'")

#         # ▼▼▼ Calcule o valor_unitario_medio e adicione ao dicionário ▼▼▼
#         valores_unitarios = [
#             row["VALOR TOTAL"] / row["AREA TOTAL"] if row["AREA TOTAL"] > 0 else 0
#             for _, row in df_ativas.iterrows()
#         ]
#         valor_unitario_medio = sum(valores_unitarios) / len([v for v in valores_unitarios if v > 0]) if valores_unitarios else 0
#         dados["dados_avaliando"]["valor_unitario_medio"] = valor_unitario_medio
#         # ▲▲▲ FIM DO BLOCO ▲▲▲

#         logger.info("📌 Aplicando Chauvenet e filtro nas amostras ativas")
#         df_filtrado, idx_excluidos, _, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_ativas)
#         logger.info(f"✅ Chauvenet concluído: {len(df_filtrado)} amostras restaram")
#         if df_filtrado.empty:
#             logger.warning("Nenhuma amostra restou após os filtros. Abortando resposta iterativa.")
#             return jsonify({"erro": "Nenhuma amostra restou após os filtros. Ative pelo menos uma amostra ou ajuste os filtros."}), 400
#         amostras_excluidas_chauvenet = [int(df_ativas.iloc[idx]["idx"]) for idx in idx_excluidos]

#         logger.info("📌 Iniciando homogeneização das amostras")
#         amostras_homog = homogeneizar_amostras(
#             df_filtrado,
#             dados["dados_avaliando"],
#             dados["fatores_do_usuario"],
#             finalidade_do_laudo=(
#                 "desapropriacao"
#                 if "desapropria" in dados["fatores_do_usuario"]["finalidade_descricao"].lower()
#                 else "servidao"
#                 if "servid" in dados["fatores_do_usuario"]["finalidade_descricao"].lower()
#                 else "mercado"
#             ),
#         )
#         logger.info("✅ Homogeneização concluída com sucesso")
       
#         #valores_unit_ativos = [a["valor_unitario"] for i, a in enumerate(amostras_homog) if i in ativos_frontend]

#         ativos_set = set(ativos_frontend)
#         valores_unit_ativos = [a["valor_unitario"] for a in amostras_homog if a.get("idx") in ativos_set]

#         array_homog = np.array([a["valor_unitario"] for a in amostras_homog], dtype=float)
#         if len(array_homog) > 1:
#             limite_inf, limite_sup = intervalo_confianca_bootstrap_mediana(array_homog, 1000, 0.80)
               
#             valor_minimo = round(limite_inf, 2)
#             valor_maximo = round(limite_sup, 2)
#             valor_medio = round(np.median(array_homog), 2)

#             amplitude_intervalo_confianca = round(((valor_maximo - valor_minimo) / valor_medio) * 100, 2)
#         else:
#             valor_minimo = valor_medio = valor_maximo = round(array_homog[0], 2)
#             amplitude_intervalo_confianca = 80  # ou outro valor padrão que desejar
#         valores_unit_ativos = [a["valor_unitario"] for i, a in enumerate(amostras_homog) if i in ativos_frontend]
#         pasta_saida = os.path.join(BASE_DIR, "static", "arquivos", f"avaliacao_{uuid}")
#         os.makedirs(pasta_saida, exist_ok=True)

#         img1 = os.path.join(pasta_saida, "grafico_aderencia_iterativo.png")
#         img2 = os.path.join(pasta_saida, "grafico_dispersao_iterativo.png")

#         amostras_chauvenet_retirou = [
#             idx for idx in ativos_frontend if idx not in df_filtrado["idx"].tolist()
#         ]

#         logger.info("📌 Gerando gráfico de dispersão iterativo")

#        # Monte os arrays sincronizados (depois da homogeneização)
#         ativos_set = set(ativos_frontend)
#         amostras_plot = [a for a in amostras_homog if a.get("idx") in ativos_set]
#         ativos_validos_idx = [a["idx"] for a in amostras_plot]
#         ativos_validos_valores = [a["valor_unitario"] for a in amostras_plot]

#         gerar_grafico_dispersao_mediana(
#             df_filtrado,
#             ativos_validos_valores,
#             img2,
#             ativos_validos_idx,
#             amostras_usuario_retirou,
#             amostras_chauvenet_retirou,
#         )

#         logger.info("✅ Gráfico dispersão gerado com sucesso")
        

#         logger.info("📌 Gerando gráfico de aderência iterativo")
#         gerar_grafico_aderencia_totais(df_filtrado, [a["valor_unitario"] for a in amostras_homog], img1)

#         logger.info("✅ Gráfico aderência gerado com sucesso")

#         resposta = {
#             "valor_minimo": valor_minimo,
#             "valor_medio": valor_medio,
#             "valor_maximo": valor_maximo,
#             "amplitude_intervalo_confianca": amplitude_intervalo_confianca,
#             "quantidade_amostras_iniciais": len(dados["amostras"]),
#             "quantidade_amostras_usuario_retirou": len(amostras_usuario_retirou),
#             "amostras_usuario_retirou": amostras_usuario_retirou,
#             "quantidade_amostras_chauvenet_retirou": len(amostras_excluidas_chauvenet),
#             "amostras_chauvenet_retirou": amostras_excluidas_chauvenet,
#             "quantidade_amostras_restantes": len(df_filtrado),
#             "grafico_dispersao_url": url_for(
#                 "static",
#                 filename=f"arquivos/avaliacao_{uuid}/grafico_dispersao_iterativo.png",
#             ),
#             "grafico_aderencia_url": url_for(
#                 "static",
#                 filename=f"arquivos/avaliacao_{uuid}/grafico_aderencia_iterativo.png",
#             ),
#         }

#         logger.info("✅ Resposta JSON pronta para envio ao frontend")
#         return jsonify(resposta)

#     except Exception as e:
#         logger.exception(f"🚨 ERRO CRÍTICO NA ROTA calcular_valores_iterativos: {e}")
#         return jsonify({"erro": f"Erro crítico interno: {str(e)}"}), 500

#         # ▼▼▼ Calcule o valor_unitario_medio e adicione ao dicionário ▼▼▼
#         valores_unitarios = [
#             row["VALOR TOTAL"] / row["AREA TOTAL"] if row["AREA TOTAL"] > 0 else 0
#             for _, row in df_ativas.iterrows()
#         ]
#         valor_unitario_medio = sum(valores_unitarios) / len([v for v in valores_unitarios if v > 0]) if valores_unitarios else 0
#         dados["dados_avaliando"]["valor_unitario_medio"] = valor_unitario_medio
#         # ▲▲▲ FIM DO BLOCO ▲▲▲

#         logger.info("📌 Aplicando Chauvenet e filtro nas amostras ativas")
#         df_filtrado, idx_excluidos, _, media, dp, menor, maior, mediana = aplicar_chauvenet_e_filtrar(df_ativas)
#         logger.info(f"✅ Chauvenet concluído: {len(df_filtrado)} amostras restaram")
#         if df_filtrado.empty:
#             logger.warning("Nenhuma amostra restou após os filtros. Abortando resposta iterativa.")
#             return jsonify({"erro": "Nenhuma amostra restou após os filtros. Ative pelo menos uma amostra ou ajuste os filtros."}), 400
#         amostras_excluidas_chauvenet = [int(df_ativas.iloc[idx]["idx"]) for idx in idx_excluidos]

#         logger.info("📌 Iniciando homogeneização das amostras")
#         amostras_homog = homogeneizar_amostras(
#             df_filtrado,
#             dados["dados_avaliando"],
#             dados["fatores_do_usuario"],
#             finalidade_do_laudo=(
#                 "desapropriacao"
#                 if "desapropria" in dados["fatores_do_usuario"]["finalidade_descricao"].lower()
#                 else "servidao"
#                 if "servid" in dados["fatores_do_usuario"]["finalidade_descricao"].lower()
#                 else "mercado"
#             ),
#         )
#         logger.info("✅ Homogeneização concluída com sucesso")
       
#         #valores_unit_ativos = [a["valor_unitario"] for i, a in enumerate(amostras_homog) if i in ativos_frontend]

#         ativos_set = set(ativos_frontend)
#         valores_unit_ativos = [a["valor_unitario"] for a in amostras_homog if a.get("idx") in ativos_set]

#         array_homog = np.array([a["valor_unitario"] for a in amostras_homog], dtype=float)
#         if len(array_homog) > 1:
#             limite_inf, limite_sup = intervalo_confianca_bootstrap_mediana(array_homog, 1000, 0.80)
               
#             valor_minimo = round(limite_inf, 2)
#             valor_maximo = round(limite_sup, 2)
#             valor_medio = round(np.median(array_homog), 2)

#             amplitude_intervalo_confianca = round(((valor_maximo - valor_minimo) / valor_medio) * 100, 2)
#         else:
#             valor_minimo = valor_medio = valor_maximo = round(array_homog[0], 2)
#             amplitude_intervalo_confianca = 80  # ou outro valor padrão que desejar
#         valores_unit_ativos = [a["valor_unitario"] for i, a in enumerate(amostras_homog) if i in ativos_frontend]
#         pasta_saida = os.path.join(BASE_DIR, "static", "arquivos", f"avaliacao_{uuid}")
#         os.makedirs(pasta_saida, exist_ok=True)

#         img1 = os.path.join(pasta_saida, "grafico_aderencia_iterativo.png")
#         img2 = os.path.join(pasta_saida, "grafico_dispersao_iterativo.png")

#         amostras_chauvenet_retirou = [
#             idx for idx in ativos_frontend if idx not in df_filtrado["idx"].tolist()
#         ]

#         logger.info("📌 Gerando gráfico de dispersão iterativo")

#        # Monte os arrays sincronizados (depois da homogeneização)
#         ativos_set = set(ativos_frontend)
#         amostras_plot = [a for a in amostras_homog if a.get("idx") in ativos_set]
#         ativos_validos_idx = [a["idx"] for a in amostras_plot]
#         ativos_validos_valores = [a["valor_unitario"] for a in amostras_plot]

#         gerar_grafico_dispersao_mediana(
#             df_filtrado,
#             ativos_validos_valores,
#             img2,
#             ativos_validos_idx,
#             amostras_usuario_retirou,
#             amostras_chauvenet_retirou,
#         )

#         logger.info("✅ Gráfico dispersão gerado com sucesso")
        

#         logger.info("📌 Gerando gráfico de aderência iterativo")
#         gerar_grafico_aderencia_totais(df_filtrado, [a["valor_unitario"] for a in amostras_homog], img1)

#         logger.info("✅ Gráfico aderência gerado com sucesso")

#         resposta = {
#             "valor_minimo": valor_minimo,
#             "valor_medio": valor_medio,
#             "valor_maximo": valor_maximo,
#             "amplitude_intervalo_confianca": amplitude_intervalo_confianca,
#             "quantidade_amostras_iniciais": len(dados["amostras"]),
#             "quantidade_amostras_usuario_retirou": len(amostras_usuario_retirou),
#             "amostras_usuario_retirou": amostras_usuario_retirou,
#             "quantidade_amostras_chauvenet_retirou": len(amostras_excluidas_chauvenet),
#             "amostras_chauvenet_retirou": amostras_excluidas_chauvenet,
#             "quantidade_amostras_restantes": len(df_filtrado),
#             "grafico_dispersao_url": url_for(
#                 "static",
#                 filename=f"arquivos/avaliacao_{uuid}/grafico_dispersao_iterativo.png",
#             ),
#             "grafico_aderencia_url": url_for(
#                 "static",
#                 filename=f"arquivos/avaliacao_{uuid}/grafico_aderencia_iterativo.png",
#             ),
#         }

#         logger.info("✅ Resposta JSON pronta para envio ao frontend")
#         return jsonify(resposta)

#     except Exception as e:
#         logger.exception(f"🚨 ERRO CRÍTICO NA ROTA calcular_valores_iterativos: {e}")
#         return jsonify({"erro": f"Erro crítico interno: {str(e)}"}), 500
@app.route("/calcular_valores_iterativos/<uuid>", methods=["POST"])
def calcular_valores_iterativos(uuid):
    """
    Recalcula métricas da tela iterativa:
    - Atualiza no JSON quais amostras estão ativas (toggle do front)
    - Aplica Chauvenet nas ativas
    - Homogeneiza e calcula IC 80% (bootstrap)
    - Gera gráficos iterativos
    Retorna JSON com números e URLs das imagens.
    """
    import os, json
    import numpy as np
    import pandas as pd
    from flask import jsonify, request, url_for

    from executaveis_avaliacao.utils_json import carregar_entrada_corrente_json
    from executaveis_avaliacao.main import (
        aplicar_chauvenet_e_filtrar,
        homogeneizar_amostras,
        intervalo_confianca_bootstrap_mediana,
        gerar_grafico_dispersao_mediana,
        gerar_grafico_aderencia_totais,
    )

    try:
        logger.info("🚀 Rota calcular_valores_iterativos iniciada")

        data = carregar_entrada_corrente_json(uuid)

        ativos_frontend = request.json.get("ativos", [])
        try:
            ativos_frontend = [int(x) for x in ativos_frontend]
        except Exception:
            return jsonify({"erro": "Formato inválido para 'ativos'"}), 400
        ativos_set = set(ativos_frontend)

        amostras         = data.get("amostras", []) or []
        dados_avaliando  = data.get("dados_avaliando", {}) or {}
        fatores_usuario  = data.get("fatores_do_usuario", {}) or {}
        cfg              = data.get("config_modelo", {}) or {}

        # Atualiza ativos no JSON
        todos_idx = [int(a.get("idx", 0) or 0) for a in amostras]
        off_set   = set(todos_idx) - ativos_set
        for a in amostras:
            a["ativo"] = int(a.get("idx", 0) or 0) in ativos_set
        cfg["amostras_desabilitadas"] = sorted(list(off_set))
        data["amostras"] = amostras
        data["config_modelo"] = cfg

        caminho_json = os.path.join(BASE_DIR, "static", "tmp", f"{uuid}_entrada_corrente.json")
        with open(caminho_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Filtra ativas
        amostras_ativas = [a for a in amostras if a.get("ativo") and float(a.get("area", 0) or 0) > 0]
        if not amostras_ativas:
            return jsonify({"erro": "Nenhuma amostra ativa."}), 400

        # Cria DF direto das amostras
        df_ativas = pd.DataFrame(amostras_ativas)
        df_ativas.rename(columns={
            "valor_total": "VALOR TOTAL",
            "area": "AREA TOTAL",
        }, inplace=True)
        if "DISTANCIA CENTRO" not in df_ativas.columns:
            df_ativas["DISTANCIA CENTRO"] = 0.0

        # valor_unitario_medio informativo
        vu = [(float(vt) / float(ar)) if float(ar or 0) > 0 else 0.0
              for vt, ar in zip(df_ativas["VALOR TOTAL"], df_ativas["AREA TOTAL"])]
        vu_validos = [v for v in vu if v > 0]
        dados_avaliando["valor_unitario_medio"] = (sum(vu_validos) / len(vu_validos)) if vu_validos else 0.0
        data["dados_avaliando"] = dados_avaliando
        with open(caminho_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Chauvenet
        df_filtrado, _, _, _, _, _, _, _ = aplicar_chauvenet_e_filtrar(df_ativas)
        if df_filtrado.empty:
            return jsonify({"erro": "Nenhuma amostra restou após os filtros."}), 400

        # Homogeneização
        finalidade = "mercado"
        fd = (fatores_usuario.get("finalidade_descricao") or "").lower()
        if "desapropria" in fd:
            finalidade = "desapropriacao"
        elif "servid" in fd:
            finalidade = "servidao"

        amostras_homog = homogeneizar_amostras(
            df_filtrado,
            dados_avaliando,
            fatores_usuario,
            finalidade_do_laudo=finalidade,
        )

        array_homog = np.array([float(a.get("valor_unitario", 0) or 0) for a in amostras_homog], dtype=float)
        array_homog = array_homog[~np.isnan(array_homog)]
        if array_homog.size == 0:
            return jsonify({"erro": "Sem valores homogêneos válidos."}), 400

        if array_homog.size == 1:
            li = ls = array_homog[0]
        else:
            bootstrap_n = int(cfg.get("parametros", {}).get("bootstrap_n", 400) or 400)
            li, ls = intervalo_confianca_bootstrap_mediana(array_homog, bootstrap_n, 0.80)

        valor_medio  = round(float(np.median(array_homog)), 2)
        valor_minimo = round(float(li), 2)
        valor_maximo = round(float(ls), 2)
        amplitude_intervalo_confianca = round(((valor_maximo - valor_minimo) / valor_medio) * 100, 2) if valor_medio > 0 else 0.0

        # Geração de gráficos
        pasta_saida = os.path.join(BASE_DIR, "static", "arquivos", f"avaliacao_{uuid}")
        os.makedirs(pasta_saida, exist_ok=True)
        img1 = os.path.join(pasta_saida, "grafico_aderencia_iterativo.png")
        img2 = os.path.join(pasta_saida, "grafico_dispersao_iterativo.png")

        gerar_grafico_aderencia_totais(df_filtrado, [a["valor_unitario"] for a in amostras_homog], img1)
        gerar_grafico_dispersao_mediana(
            df_filtrado,
            [a["valor_unitario"] for a in amostras_homog],
            img2,
            [a["idx"] for a in amostras_homog],
            sorted(list(off_set)),
            sorted(set(df_ativas["idx"]) - set(df_filtrado["idx"])),
        )

        resposta = {
            "valor_minimo": valor_minimo,
            "valor_medio": valor_medio,
            "valor_maximo": valor_maximo,
            "amplitude_intervalo_confianca": amplitude_intervalo_confianca,
            "quantidade_amostras_iniciais": len(amostras),
            "quantidade_amostras_usuario_retirou": len(off_set),
            "amostras_usuario_retirou": sorted(list(off_set)),
            "quantidade_amostras_chauvenet_retirou": len(set(df_ativas["idx"]) - set(df_filtrado["idx"])),
            "amostras_chauvenet_retirou": sorted(set(df_ativas["idx"]) - set(df_filtrado["idx"])),
            "quantidade_amostras_restantes": int(df_filtrado.shape[0]),
            "grafico_dispersao_url": url_for("static", filename=f"arquivos/avaliacao_{uuid}/grafico_dispersao_iterativo.png"),
            "grafico_aderencia_url": url_for("static", filename=f"arquivos/avaliacao_{uuid}/grafico_aderencia_iterativo.png"),
        }
        return jsonify(resposta)

    except FileNotFoundError:
        return jsonify({"erro": "Arquivo de entrada não encontrado."}), 400
    except Exception as e:
        logger.exception(f"🚨 ERRO CRÍTICO NA ROTA calcular_valores_iterativos: {e}")
        return jsonify({"erro": f"Erro crítico interno: {str(e)}"}), 500


 
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)


@app.route('/debug_rotas')
def debug_rotas():
    from flask import Response
    rotas = [str(rule) for rule in app.url_map.iter_rules()]
    return Response('<br>'.join(rotas), mimetype='text/html')

#finalizando#


