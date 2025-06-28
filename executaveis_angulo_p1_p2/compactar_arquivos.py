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

        # Padrão para arquivos DXF e DOCX finais (unificados)
        padrao_dxf_final = os.path.join(diretorio, f"{uuid_str}_{tipo}_*_FINAL.dxf")
        padrao_docx_final = os.path.join(diretorio, f"{uuid_str}_{tipo}_*_FINAL.docx")

        # Padrões para arquivos Excel ABERTA e FECHADA
        padrao_excel_aberta = os.path.join(diretorio, f"{uuid_str}_ABERTA_{tipo}_*.xlsx")
        padrao_excel_fechada = os.path.join(diretorio, f"{uuid_str}_FECHADA_{tipo}_*.xlsx")

        arquivo_dxf_final = glob.glob(padrao_dxf_final)
        arquivo_docx_final = glob.glob(padrao_docx_final)
        arquivo_excel_aberta = glob.glob(padrao_excel_aberta)
        arquivo_excel_fechada = glob.glob(padrao_excel_fechada)

        print(f"   - DXF FINAL encontrados: {len(arquivo_dxf_final)}")
        print(f"   - DOCX FINAL encontrados: {len(arquivo_docx_final)}")
        print(f"   - XLSX ABERTA encontrados: {len(arquivo_excel_aberta)}")
        print(f"   - XLSX FECHADA encontrados: {len(arquivo_excel_fechada)}")

        if arquivo_dxf_final and arquivo_docx_final and arquivo_excel_aberta and arquivo_excel_fechada:
            nome_zip = f"{uuid_str}_{cidade_formatada}_{tipo}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivo_dxf_final[0], os.path.basename(arquivo_dxf_final[0]))
                    zipf.write(arquivo_docx_final[0], os.path.basename(arquivo_docx_final[0]))
                    zipf.write(arquivo_excel_aberta[0], os.path.basename(arquivo_excel_aberta[0]))
                    zipf.write(arquivo_excel_fechada[0], os.path.basename(arquivo_excel_fechada[0]))

                print(f"🗜️ ZIP salvo em: {caminho_zip}")
                logger.info(f"ZIP criado: {caminho_zip}")

            except Exception as e:
                print(f"❌ Erro ao criar ZIP: {e}")
                logger.error(f"Erro ao criar ZIP {caminho_zip}: {e}")
        else:
            print(f"⚠️ Arquivos incompletos ou não encontrados para o tipo {tipo}")
            logger.warning(
                f"Incompleto: {tipo} - "
                f"DXF={bool(arquivo_dxf_final)}, DOCX={bool(arquivo_docx_final)}, "
                f"XLSX ABERTA={bool(arquivo_excel_aberta)}, XLSX FECHADA={bool(arquivo_excel_fechada)}"
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
