import os
import glob
import zipfile
import re
import logging
from datetime import datetime
import shutil
import json

# Defina BASE_DIR de forma segura para execução local
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not any(isinstance(h, logging.FileHandler) and getattr(h, "baseFilename", None) == log_file for h in logger.handlers):
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    logger.addHandler(file_handler)

# cidade segura (sem espaços, barras, acentos etc.)
def _safe_city(s):
    s = (s or "CIDADE")
    s = re.sub(r'[^A-Za-z0-9_.-]+', '_', s)  # troca qualquer coisa não segura por _
    s = s.strip('._-')
    return s or "CIDADE"

# matrícula segura (remove lixo e pontuação final duplicada)
def _safe_mat(s):
    s = (s or "").strip()
    s = s.replace(" ", "").replace(",", ".").replace("__", "_")
    s = re.sub(r'[^0-9A-Za-z._-]+', '', s)   # mantém só [0-9A-Za-z._-]
    s = s.strip('._-')                        # tira pontuação nas pontas
    return s

# def montar_pacote_zip(diretorio, cidade):
#     print("\n📦 [compactar] Iniciando montagem dos pacotes ZIP")
#     logger.info("Iniciando montagem dos pacotes ZIP")

#     created_zips = []  # nomes (base) dos zips gerados nesta execução (em diretorio)
#     uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(diretorio)))

#     tipos = ["ETE", "REM", "SER", "ACE"]

#     print("\n📁 [DEBUG] Listando todos os arquivos no diretório:")
#     try:
#         for arquivo in os.listdir(diretorio):
#             print("🗂️", arquivo)
#     except FileNotFoundError:
#         logger.error(f"Diretório não encontrado: {diretorio}")
#         return

#     for tipo in tipos:
#         print(f"\n🔍 Buscando arquivos do tipo: {tipo}")
#         logger.info(f"Buscando arquivos do tipo: {tipo}")

#         print(f"[DEBUG compactar] UUID identificado: {uuid_prefix}")

#         padrao_dxf = os.path.join(diretorio, f"*{tipo}*.dxf")
#         padrao_docx = os.path.join(diretorio, f"*{tipo}*.docx")
#         padrao_excel = os.path.join(diretorio, f"*{tipo}*.xlsx")

#         print(f"🧭 [DEBUG] Padrões de busca:")
#         print(f"   - DXF  : {padrao_dxf}")
#         print(f"   - DOCX : {padrao_docx}")
#         print(f"   - XLSX : {padrao_excel}")

#         arquivos_dxf = glob.glob(padrao_dxf)
#         arquivos_docx = glob.glob(padrao_docx)
#         arquivos_excel = glob.glob(padrao_excel)

#         print(f"   - DXF encontrados: {len(arquivos_dxf)}")
#         print(f"   - DOCX encontrados: {len(arquivos_docx)}")
#         print(f"   - XLSX encontrados: {len(arquivos_excel)}")

#         logger.info(f"DXF={len(arquivos_dxf)} | DOCX={len(arquivos_docx)} | XLSX={len(arquivos_excel)}")

#         # Extrai matrículas observadas nos nomes de arquivo desse tipo
#         matriculas = set()
#         for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
#             nome_arquivo = os.path.basename(arq)
#             match = re.search(rf"{tipo}.*?(\d{{2,6}}[.,_]?\d{{0,3}})", nome_arquivo, re.IGNORECASE)
#             if match:
#                 matricula = match.group(1).replace(",", ".").replace(" ", "")
#                 if "." not in matricula and len(matricula) > 3:
#                     matricula = f"{matricula[:-3]}.{matricula[-3:]}"
#                 matriculas.add(matricula)

#         for matricula in matriculas:
#             print(f"\n🔢 Processando matrícula: {matricula}")
#             logger.info(f"Processando matrícula: {matricula}")

#             arq_dxf = [a for a in arquivos_dxf if matricula in a]
#             arq_docx = [a for a in arquivos_docx if matricula in a]
#             arq_excel = [a for a in arquivos_excel if matricula in a]

#             if arq_dxf and arq_docx and arq_excel:
#                 cidade_sanitizada = _safe_city(cidade)
#                 matricula_sanit   = _safe_mat(matricula)
#                 #cidade_sanitizada = (cidade or "CIDADE").replace(" ", "_")
#                 zip_base = f"{cidade_sanitizada}_{tipo}_{matricula_sanit}.zip"
#                 #nome_zip = os.path.join(diretorio, zip_base)
#                 # Depois (inclui UUID):
#                 nome_zip = f"{uuid_exec}_{cidade}_{tipo}_{matricula}.zip"  # ex.: 1d0ca532_CHUI_ETE_32.681.zip
#                 logger.info(f"[ZIP] nome_zip={zip_base} (cidade='{cidade}', matricula='{matricula}')")

#                 STATIC_ZIP_DIR = os.path.join(BASE_DIR, 'static', 'arquivos')
#                 os.makedirs(STATIC_ZIP_DIR, exist_ok=True)
#                 caminho_debug_zip = os.path.join(STATIC_ZIP_DIR, f"{uuid_prefix}_{os.path.basename(nome_zip)}")

#                 try:
#                     with zipfile.ZipFile(nome_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
#                         zipf.write(arq_docx[0], arcname=os.path.basename(arq_docx[0]))
#                         zipf.write(arq_dxf[0], arcname=os.path.basename(arq_dxf[0]))
#                         zipf.write(arq_excel[0], arcname=os.path.basename(arq_excel[0]))

#                     # registra o zip gerado (nome base dentro do CONCLUIDO)
#                     created_zips.append(os.path.basename(nome_zip))

#                     # cópia para a área pública (download)
#                     try:
#                         shutil.copy2(nome_zip, caminho_debug_zip)
                     
#                         print(f"🪣 ZIP também copiado para: {caminho_debug_zip}")
#                     except Exception as e_copy:
#                         logger.warning(f"Falha ao copiar ZIP para público: {e_copy}")

#                     print(f"✅ ZIP criado com sucesso: {nome_zip}")
#                     logger.info(f"ZIP criado: {nome_zip} e (tentativa de) cópia para: {caminho_debug_zip}")
#                 except Exception as e:
#                     logger.exception(f"Erro ao criar ZIP {nome_zip}")
#                     print(f"❌ Erro ao criar ZIP: {e}")
#             else:
#                 logger.info(f"Arquivos insuficientes para {tipo} - matrícula {matricula}: "
#                             f"DXF={len(arq_dxf)} DOCX={len(arq_docx)} XLSX={len(arq_excel)}")

#     # Escreve manifesto com os zips desta execução no CONCLUIDO
#     try:
#         run_json = os.path.join(diretorio, "RUN.json")
#         with open(run_json, "w", encoding="utf-8") as f:
#             json.dump({"zip_files": created_zips}, f, ensure_ascii=False)
#         logger.info(f"[RUN] Manifesto salvo: {run_json} | zip_files={created_zips}")
#     except Exception as e:
#         logger.warning(f"[RUN] Falha ao salvar RUN.json: {e}")


def montar_pacote_zip(diretorio, cidade):
    print("\n📦 [compactar] Iniciando montagem dos pacotes ZIP")
    logger.info("Iniciando montagem dos pacotes ZIP")

    created_zips = []  # nomes dos zips gerados nesta execução (apenas o filename)
    uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(diretorio)))
    # >>> PATCH: padroniza o UUID curto e seguro
    uuid_exec = re.sub(r'[^0-9a-fA-F]', '', uuid_prefix)[:8] or uuid_prefix
    # <<< PATCH

    tipos = ["ETE", "REM", "SER", "ACE"]

    print("\n📁 [DEBUG] Listando todos os arquivos no diretório:")
    try:
        for arquivo in os.listdir(diretorio):
            print("🗂️", arquivo)
    except FileNotFoundError:
        logger.error(f"Diretório não encontrado: {diretorio}")
        return

    for tipo in tipos:
        print(f"\n🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        print(f"[DEBUG compactar] UUID identificado: {uuid_exec}")

        padrao_dxf = os.path.join(diretorio, f"*{tipo}*.dxf")
        padrao_docx = os.path.join(diretorio, f"*{tipo}*.docx")
        padrao_excel = os.path.join(diretorio, f"*{tipo}*.xlsx")

        print(f"🧭 [DEBUG] Padrões de busca:")
        print(f"   - DXF  : {padrao_dxf}")
        print(f"   - DOCX : {padrao_docx}")
        print(f"   - XLSX : {padrao_excel}")

        arquivos_dxf = glob.glob(padrao_dxf)
        arquivos_docx = glob.glob(padrao_docx)
        arquivos_excel = glob.glob(padrao_excel)

        print(f"   - DXF encontrados: {len(arquivos_dxf)}")
        print(f"   - DOCX encontrados: {len(arquivos_docx)}")
        print(f"   - XLSX encontrados: {len(arquivos_excel)}")

        logger.info(f"DXF={len(arquivos_dxf)} | DOCX={len(arquivos_docx)} | XLSX={len(arquivos_excel)}")

        # Extrai matrículas observadas nos nomes de arquivo desse tipo
        matriculas = set()
        for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
            nome_arquivo = os.path.basename(arq)
            match = re.search(rf"{tipo}.*?(\d{{2,6}}[.,_]?\d{{0,3}})", nome_arquivo, re.IGNORECASE)
            if match:
                matricula = match.group(1).replace(",", ".").replace(" ", "")
                if "." not in matricula and len(matricula) > 3:
                    matricula = f"{matricula[:-3]}.{matricula[-3:]}"
                matriculas.add(matricula)

        for matricula in matriculas:
            print(f"\n🔢 Processando matrícula: {matricula}")
            logger.info(f"Processando matrícula: {matricula}")

            arq_dxf = [a for a in arquivos_dxf if matricula in a]
            arq_docx = [a for a in arquivos_docx if matricula in a]
            arq_excel = [a for a in arquivos_excel if matricula in a]

            if arq_dxf and arq_docx and arq_excel:
                # >>> PATCH: nomes 100% padronizados (mesmo nome em CONCLUIDO e em static/arquivos)
                cidade_sanitizada = _safe_city(cidade)
                matricula_sanit   = _safe_mat(matricula)
                zip_filename      = f"{uuid_exec}_{cidade_sanitizada}_{tipo}_{matricula_sanit}.zip"
                zip_path_conc     = os.path.join(diretorio, zip_filename)

                logger.info(f"[ZIP] nome_zip={zip_filename} (cidade='{cidade}', matricula='{matricula}')")

                STATIC_ZIP_DIR = os.path.join(BASE_DIR, 'static', 'arquivos')
                os.makedirs(STATIC_ZIP_DIR, exist_ok=True)
                dest_static = os.path.join(STATIC_ZIP_DIR, zip_filename)
                # <<< PATCH

                try:
                    # cria o ZIP **dentro do CONCLUIDO** com o nome final
                    with zipfile.ZipFile(zip_path_conc, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
                        zipf.write(arq_docx[0], arcname=os.path.basename(arq_docx[0]))
                        zipf.write(arq_dxf[0],  arcname=os.path.basename(arq_dxf[0]))
                        zipf.write(arq_excel[0],arcname=os.path.basename(arq_excel[0]))

                    # registra o zip gerado
                    created_zips.append(zip_filename)

                    # cópia para a área pública (download) com **o MESMO** nome
                    try:
                        shutil.copy2(zip_path_conc, dest_static)
                        print(f"🪣 ZIP também copiado para: {dest_static}")
                    except Exception as e_copy:
                        logger.warning(f"Falha ao copiar ZIP para público: {e_copy}")

                    print(f"✅ ZIP criado com sucesso: {zip_path_conc}")
                    logger.info(f"ZIP criado: {zip_path_conc} | copiado para: {dest_static}")
                except Exception as e:
                    logger.exception(f"Erro ao criar ZIP {zip_path_conc}")
                    print(f"❌ Erro ao criar ZIP: {e}")
            else:
                logger.info(f"Arquivos insuficientes para {tipo} - matrícula {matricula}: "
                            f"DXF={len(arq_dxf)} DOCX={len(arq_docx)} XLSX={len(arq_excel)}")

    # >>> PATCH: manifesto com mais contexto e nomes finais (com UUID)
    try:
        run_json = os.path.join(diretorio, "RUN.json")
        with open(run_json, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "uuid": uuid_exec,
                    "cidade": _safe_city(cidade),
                    "zip_files": created_zips,   # filenames finais, idênticos em CONCLUIDO e static/arquivos
                    "concluido_dir": diretorio
                },
                f,
                ensure_ascii=False,
                indent=2
            )
        logger.info(f"[RUN] Manifesto salvo: {run_json} | zip_files={created_zips}")
    except Exception as e:
        logger.warning(f"[RUN] Falha ao salvar RUN.json: {e}")
    # <<< PATCH


def main_compactar_arquivos(diretorio_concluido, cidade_formatada):
    print(f"\n📦 Iniciando compactação no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade_formatada)
    print("✅ Compactação finalizada")
    logger.info("Compactação finalizada")
