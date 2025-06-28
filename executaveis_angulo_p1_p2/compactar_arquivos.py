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

def montar_pacote_zip(diretorio, cidade_formatada, uuid_str):

    print(f"\n📦 Compactando arquivos no diretório: {diretorio}")
    logger.info(f"Iniciando montagem dos pacotes ZIP")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        print(f"🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        arquivos_dxf = glob.glob(os.path.join(diretorio, f"{uuid_str}_{tipo}*_FINAL.dxf"))
        arquivos_docx = glob.glob(os.path.join(diretorio, f"{uuid_str}_{tipo}*_FINAL.docx"))
        arquivos_excel_aberta = glob.glob(os.path.join(diretorio, f"{uuid_str}_ABERTA_{tipo}*.xlsx"))
        arquivos_excel_fechada = glob.glob(os.path.join(diretorio, f"{uuid_str}_FECHADA_{tipo}*.xlsx"))

        print(f"   - DXF FINAL encontrados: {len(arquivos_dxf)}")
        print(f"   - DOCX FINAL encontrados: {len(arquivos_docx)}")
        print(f"   - XLSX ABERTA encontrados: {len(arquivos_excel_aberta)}")
        print(f"   - XLSX FECHADA encontrados: {len(arquivos_excel_fechada)}")

        if arquivos_dxf and arquivos_docx and arquivos_excel_aberta and arquivos_excel_fechada:
            nome_zip = f"{uuid_str}_{cidade_formatada}_{tipo}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivos_dxf[0], os.path.basename(arquivos_dxf[0]))
                    zipf.write(arquivos_docx[0], os.path.basename(arquivos_docx[0]))
                    zipf.write(arquivos_excel_aberta[0], os.path.basename(arquivos_excel_aberta[0]))
                    zipf.write(arquivos_excel_fechada[0], os.path.basename(arquivos_excel_fechada[0]))

                print(f"🗜️ ZIP salvo em: {caminho_zip}")
                logger.info(f"ZIP criado: {caminho_zip}")

            except Exception as e:
                print(f"❌ Erro ao criar ZIP: {e}")
                logger.error(f"Erro ao criar ZIP {caminho_zip}: {e}")
        else:
            print(f"⚠️ Arquivos incompletos ou não encontrados para o tipo {tipo}")
            logger.warning(
                f"Incompleto: {tipo} - "
                f"DXF={bool(arquivos_dxf)}, DOCX={bool(arquivos_docx)}, "
                f"XLSX ABERTA={bool(arquivos_excel_aberta)}, XLSX FECHADA={bool(arquivos_excel_fechada)}"
            )




def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):

    print(f"\n📦 Compactando arquivos no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    print("✅ Compactação finalizada")
    logger.info("Compactação finalizada")

if __name__ == "__main__":
    import argparse
    TMP_CONCLUIDO = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
    parser = argparse.ArgumentParser(description="Compacta arquivos gerados.")
    parser.add_argument('--diretorio', default=TMP_CONCLUIDO, help="Diretório de saída.")
    parser.add_argument('--cidade', default="Cidade", help="Nome da cidade.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio, args.cidade)
