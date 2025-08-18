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
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def montar_pacote_zip(diretorio, cidade):
    print("\n📦 [compactar] Iniciando montagem dos pacotes ZIP")
    logger.info("Iniciando montagem dos pacotes ZIP")

    created_zips = []  # nomes (base) dos zips gerados nesta execução (em diretorio)
    uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(diretorio)))


    tipos = ["ETE", "REM", "SER", "ACE"]

    print("\n📁 [DEBUG] Listando todos os arquivos no diretório:")
    for arquivo in os.listdir(diretorio):
        print("🗂️", arquivo)

    for tipo in tipos:
        print(f"\n🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        print(f"[DEBUG compactar] UUID identificado: {uuid_prefix}")


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
                cidade_sanitizada = cidade.replace(" ", "_")
                nome_zip = os.path.join(diretorio, f"{cidade_sanitizada}_{tipo}_{matricula}.zip")
                
                STATIC_ZIP_DIR = os.path.join(BASE_DIR, 'static', 'arquivos')
                os.makedirs(STATIC_ZIP_DIR, exist_ok=True)
                caminho_debug_zip = os.path.join(STATIC_ZIP_DIR, f"{uuid_prefix}_{os.path.basename(nome_zip)}")


                try:
                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(arq_docx[0], arcname=os.path.basename(arq_docx[0]))
                        zipf.write(arq_dxf[0], arcname=os.path.basename(arq_dxf[0]))
                        zipf.write(arq_excel[0], arcname=os.path.basename(arq_excel[0]))
                    
                    # registra o zip gerado (nome base dentro do CONCLUIDO)
                    created_zips.append(os.path.basename(nome_zip))


                    print(f"✅ ZIP criado com sucesso: {nome_zip}")
                    logger.info(f"ZIP criado: {nome_zip} e copiado para: {caminho_debug_zip}")
                except Exception as e:
                    logger.exception(f"Erro ao criar ZIP {nome_zip}")
                    print(f"❌ Erro ao criar ZIP: {e}")
    # Escreve manifesto com os zips desta execução no CONCLUIDO
    try:
        run_json = os.path.join(diretorio, "RUN.json")
        with open(run_json, "w", encoding="utf-8") as f:
            json.dump({"zip_files": created_zips}, f, ensure_ascii=False)
        logger.info(f"[RUN] Manifesto salvo: {run_json} | zip_files={created_zips}")
    except Exception as e:
        logger.warning(f"[RUN] Falha ao salvar RUN.json: {e}")


def main_compactar_arquivos(diretorio_concluido, cidade_formatada):
    print(f"\n📦 Iniciando compactação no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade_formatada)
    print("✅ Compactação finalizada")
    logger.info("Compactação finalizada")
