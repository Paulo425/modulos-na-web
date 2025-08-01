import os
import glob
import zipfile
import logging
from datetime import datetime

# Configuração inicial do diretório e log
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
    logger.info(f"Iniciando montagem dos pacotes ZIP em {diretorio}")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        padrao_dxf = os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.dxf")
        padrao_docx = os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.docx")
        padrao_excel = os.path.join(diretorio, f"{uuid_str}_*_{tipo}_*.xlsx")

        arquivo_dxf = glob.glob(padrao_dxf)
        arquivo_docx = glob.glob(padrao_docx)
        arquivo_excel = glob.glob(padrao_excel)

        logger.info(
            f"Encontrados: DXF={len(arquivo_dxf)}, DOCX={len(arquivo_docx)}, XLSX={len(arquivo_excel)}"
        )

        if arquivo_dxf and arquivo_docx and arquivo_excel:
            nome_zip = f"{uuid_str}_{cidade_formatada}_{tipo}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivo_dxf[0], os.path.basename(arquivo_dxf[0]))
                    zipf.write(arquivo_docx[0], os.path.basename(arquivo_docx[0]))
                    zipf.write(arquivo_excel[0], os.path.basename(arquivo_excel[0]))

                logger.info(f"ZIP criado com sucesso: {caminho_zip}")

            except Exception as e:
                logger.error(f"Erro ao criar ZIP {caminho_zip}: {e}")

        else:
            logger.warning(
                f"Incompleto: {tipo} - DXF={bool(arquivo_dxf)}, DOCX={bool(arquivo_docx)}, XLSX={bool(arquivo_excel)}"
            )

def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    logger.info("Compactação finalizada")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Compacta arquivos gerados.")
    parser.add_argument('--diretorio', required=True, help="Diretório dos arquivos concluídos.")
    parser.add_argument('--cidade', required=True, help="Nome formatado da cidade.")
    parser.add_argument('--uuid', required=True, help="UUID da execução atual.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio, args.cidade, args.uuid)
