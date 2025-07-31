import os
import glob
import zipfile
import logging
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def montar_pacote_zip(diretorio, cidade_formatada, uuid_str):
    logger.info(f"[AZIMUTE_AZ] Iniciando compactação no diretório: {diretorio}")
    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        logger.info(f"Buscando arquivos para o tipo: {tipo}")

        arquivo_dxf = glob.glob(os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.dxf"))
        arquivo_docx = glob.glob(os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.docx"))
        arquivo_excel = glob.glob(os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.xlsx"))

        logger.info(f"Arquivos encontrados: DXF={len(arquivo_dxf)}, DOCX={len(arquivo_docx)}, XLSX={len(arquivo_excel)}")

        if arquivo_dxf and arquivo_docx and arquivo_excel:
            nome_zip = f"{uuid_str}_{cidade_formatada}_{tipo}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivo_dxf[0], os.path.basename(arquivo_dxf[0]))
                    zipf.write(arquivo_docx[0], os.path.basename(arquivo_docx[0]))
                    zipf.write(arquivo_excel[0], os.path.basename(arquivo_excel[0]))

                logger.info(f"✅ ZIP criado: {caminho_zip}")

            except Exception as e:
                logger.error(f"❌ Erro ao criar ZIP {caminho_zip}: {e}")

        else:
            logger.warning(
                f"⚠️ Arquivos incompletos para {tipo}: "
                f"DXF={bool(arquivo_dxf)}, DOCX={bool(arquivo_docx)}, XLSX={bool(arquivo_excel)}"
            )

def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    logger.info("Compactação concluída com sucesso.")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Compacta arquivos gerados.")
    parser.add_argument('--diretorio', required=True, help="Diretório dos arquivos concluídos.")
    parser.add_argument('--cidade', required=True, help="Nome formatado da cidade.")
    parser.add_argument('--uuid', required=True, help="UUID da execução atual.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio, args.cidade, args.uuid)
