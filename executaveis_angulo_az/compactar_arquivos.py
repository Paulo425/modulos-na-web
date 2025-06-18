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

def montar_pacote_zip(diretorio, cidade_formatada):
    print(f"\n📦 Compactando arquivos no diretório: {diretorio}")
    logger.info(f"[AZIMUTE_AZ] Iniciando montagem dos pacotes ZIP")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        print(f"🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        padrao_dxf = os.path.join(diretorio, f"*{tipo}*Memorial*.dxf")
        padrao_docx = os.path.join(diretorio, f"*{tipo}*Memorial*.docx")
        padrao_excel = os.path.join(diretorio, f"*{tipo}*Memorial*.xlsx")


        arquivo_dxf = glob.glob(padrao_dxf)
        arquivo_docx = glob.glob(padrao_docx)
        arquivo_excel = glob.glob(padrao_excel)

        print(f"   - DXF encontrados: {len(arquivo_dxf)}")
        print(f"   - DOCX encontrados: {len(arquivo_docx)}")
        print(f"   - XLSX encontrados: {len(arquivo_excel)}")

        if arquivo_dxf and arquivo_docx and arquivo_excel:
            base_nome = os.path.splitext(os.path.basename(arquivo_dxf[0]))[0]
            partes = base_nome.split("_", 1)
            sufixo_identificador = partes[1] if len(partes) > 1 else partes[0]

            nome_zip = f"{cidade_formatada}_{tipo}_{sufixo_identificador}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivo_dxf[0], os.path.basename(arquivo_dxf[0]))
                    zipf.write(arquivo_docx[0], os.path.basename(arquivo_docx[0]))
                    zipf.write(arquivo_excel[0], os.path.basename(arquivo_excel[0]))

                print(f"🗜️ ZIP salvo em: {caminho_zip}")
                logger.info(f"ZIP criado: {caminho_zip}")

            except Exception as e:
                print(f"❌ Erro ao criar ZIP: {e}")
                logger.error(f"Erro ao criar ZIP {caminho_zip}: {e}")
        else:
            print(f"⚠️ Arquivos incompletos ou não encontrados para o tipo {tipo}")
            logger.warning(f"Incompleto: {tipo} - DXF={bool(arquivo_dxf)}, DOCX={bool(arquivo_docx)}, XLSX={bool(arquivo_excel)}")


def main_compactar_arquivos(diretorio_concluido, cidade):
    print(f"\n📦 Compactando arquivos no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade)
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
