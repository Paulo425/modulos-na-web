import os
import glob
import zipfile
import re
import logging
from datetime import datetime
from shutil import copyfile

# Diretórios e setup de log
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# Configurar logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def montar_pacote_zip(diretorio, cidade):
    print("\n📦 [compactar] Iniciando montagem dos pacotes ZIP")
    logger.info("Iniciando montagem dos pacotes ZIP")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        print(f"🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        arquivos_dxf = glob.glob(os.path.join(diretorio, f"*{tipo}*.dxf"))
        arquivos_docx = glob.glob(os.path.join(diretorio, f"*{tipo}*.docx"))
        arquivos_excel = glob.glob(os.path.join(diretorio, f"*{tipo}*.xlsx"))

        print(f"   - DXF encontrados: {len(arquivos_dxf)}")
        print(f"   - DOCX encontrados: {len(arquivos_docx)}")
        print(f"   - XLSX encontrados: {len(arquivos_excel)}")

        logger.info(f"DXF={len(arquivos_dxf)} | DOCX={len(arquivos_docx)} | XLSX={len(arquivos_excel)}")

        matriculas = set()
        for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
            nome_arquivo = os.path.basename(arq)
            match = re.search(r"([0-9]+)[., ]?([0-9]{3})", nome_arquivo)
            if match:
                matricula = f"{match.group(1)}.{match.group(2)}"
                matriculas.add(matricula)
                if not "." in matricula:
                    if len(matricula) > 2:
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

                STATIC_ZIP_DIR = os.path.join(BASE_DIR, 'static', 'zips')
                os.makedirs(STATIC_ZIP_DIR, exist_ok=True)
                caminho_debug_zip = os.path.join(STATIC_ZIP_DIR, os.path.basename(nome_zip))

                try:
                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(arq_dxf[0], os.path.basename(arq_dxf[0]))
                        zipf.write(arq_docx[0], os.path.basename(arq_docx[0]))
                        zipf.write(arq_excel[0], os.path.basename(arq_excel[0]))
                    copyfile(nome_zip, caminho_debug_zip)

                    print(f"✅ ZIP criado com sucesso: {nome_zip}")
                    logger.info(f"ZIP criado: {nome_zip} e copiado para: {caminho_debug_zip}")
                except Exception as e:
                    print(f"❌ Erro ao criar ZIP: {e}")
                    logger.error(f"Erro ao criar ZIP {nome_zip}: {e}")
            else:
                print(f"⚠️ Arquivos incompletos para {tipo}, matrícula {matricula}")
                logger.warning(f"Incompleto: {tipo} | matrícula {matricula} | DXF={bool(arq_dxf)}, DOCX={bool(arq_docx)}, XLSX={bool(arq_excel)}")

def main_compactar_arquivos(diretorio_concluido, cidade_formatada):
    print(f"\n📦 Iniciando compactação no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade_formatada)
    print("✅ Compactação finalizada")
    logger.info("Compactação finalizada")

if __name__ == "__main__":
    import argparse

    TMP_CONCLUIDO = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
    parser = argparse.ArgumentParser(description="Compacta arquivos gerados em ZIP.")
    parser.add_argument('--diretorio', default=TMP_CONCLUIDO, help="Diretório com os arquivos a compactar.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio)
