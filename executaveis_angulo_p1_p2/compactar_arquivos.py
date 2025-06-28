import os
import glob
import zipfile
import logging
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
)
logger = logging.getLogger(__name__)

def montar_pacote_zip(diretorio, cidade_formatada, uuid_str):
    logger.info(f"📦 Iniciando montagem dos pacotes ZIP no diretório: {diretorio}")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        logger.info(f"🔍 Processando tipo: {tipo}")

        padrao_dxf = f"*{tipo}*_FINAL.dxf"
        padrao_docx = f"*{tipo}*_FINAL.docx"
        padrao_excel_aberta = f"*ABERTA*{tipo}*.xlsx"
        padrao_excel_fechada = f"*FECHADA*{tipo}*.xlsx"

        arquivos_dxf = glob.glob(os.path.join(diretorio, padrao_dxf))
        arquivos_docx = glob.glob(os.path.join(diretorio, padrao_docx))
        arquivos_excel_aberta = glob.glob(os.path.join(diretorio, padrao_excel_aberta))
        arquivos_excel_fechada = glob.glob(os.path.join(diretorio, padrao_excel_fechada))

        logger.info(f"📂 Arquivos encontrados para '{tipo}':")
        logger.info(f"  DXF FINAL: {arquivos_dxf}")
        logger.info(f"  DOCX FINAL: {arquivos_docx}")
        logger.info(f"  XLSX ABERTA: {arquivos_excel_aberta}")
        logger.info(f"  XLSX FECHADA: {arquivos_excel_fechada}")

        if arquivos_dxf and arquivos_docx and arquivos_excel_aberta and arquivos_excel_fechada:
            nome_zip = f"{uuid_str}_{cidade_formatada}_{tipo}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            logger.info(f"📦 Criando ZIP: {caminho_zip}")

            try:
                with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                    zipf.write(arquivos_dxf[0], os.path.basename(arquivos_dxf[0]))
                    zipf.write(arquivos_docx[0], os.path.basename(arquivos_docx[0]))
                    zipf.write(arquivos_excel_aberta[0], os.path.basename(arquivos_excel_aberta[0]))
                    zipf.write(arquivos_excel_fechada[0], os.path.basename(arquivos_excel_fechada[0]))

                logger.info(f"✅ ZIP criado com sucesso: {caminho_zip}")

            except Exception as e:
                logger.error(f"❌ Erro ao criar ZIP {caminho_zip}: {e}")
        else:
            logger.warning(f"⚠️ Arquivos incompletos para {tipo}: DXF={bool(arquivos_dxf)}, DOCX={bool(arquivos_docx)}, XLSX ABERTA={bool(arquivos_excel_aberta)}, XLSX FECHADA={bool(arquivos_excel_fechada)}")

def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):
    logger.info(f"📦 Iniciando compactação final no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    logger.info("✅ Compactação finalizada")

if __name__ == "__main__":
    import argparse
    TMP_CONCLUIDO = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
    parser = argparse.ArgumentParser(description="Compacta arquivos gerados.")
    parser.add_argument('--diretorio', default=TMP_CONCLUIDO, help="Diretório de saída.")
    parser.add_argument('--cidade', default="Cidade", help="Nome da cidade.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio, args.cidade, "UUID_TESTE")
