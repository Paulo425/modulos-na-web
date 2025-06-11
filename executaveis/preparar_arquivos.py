# preparar_arquivos.py

import os
import shutil
import pandas as pd
import logging
from datetime import datetime

# Diretórios e logger
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

log_file = os.path.join(LOG_DIR, f'preparo_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def main_preparo_arquivos(diretorio_base, cidade, caminho_excel, caminho_dxf):
    TMP_DIR = os.path.join(BASE_DIR, 'tmp')
    RECEBIDO = os.path.join(TMP_DIR, 'RECEBIDO')
    PREPARADO = os.path.join(TMP_DIR, 'PREPARADO')
    CONCLUIDO = os.path.join(TMP_DIR, 'CONCLUIDO')

    for pasta in [RECEBIDO, PREPARADO, CONCLUIDO]:
        os.makedirs(pasta, exist_ok=True)

    nome_excel = os.path.basename(caminho_excel)
    nome_dxf = os.path.basename(caminho_dxf)

    destino_excel = os.path.join(RECEBIDO, nome_excel)
    destino_dxf = os.path.join(RECEBIDO, nome_dxf)

    try:
        shutil.copy(caminho_excel, destino_excel)
        print(f"✅ Excel copiado para: {destino_excel}")
        logger.info(f"Excel copiado para: {destino_excel}")
    except Exception as e:
        print(f"❌ Erro ao copiar arquivo Excel: {e}")
        logger.error(f"Erro ao copiar arquivo Excel: {e}")
        return None

    try:
        shutil.copy(caminho_dxf, destino_dxf)
        print(f"✅ DXF copiado para: {destino_dxf}")
        logger.info(f"DXF copiado para: {destino_dxf}")
    except Exception as e:
        print(f"❌ Erro ao copiar arquivo DXF: {e}")
        logger.error(f"Erro ao copiar arquivo DXF: {e}")
        return None

    try:
        df = pd.read_excel(destino_excel, sheet_name=None)
        for nome_aba, tabela in df.items():
            nome_arquivo = f"{nome_aba}_PREPARADO.xlsx"
            caminho_saida = os.path.join(PREPARADO, nome_arquivo)
            tabela.to_excel(caminho_saida, index=False)
            print(f"✅ Planilha '{nome_aba}' salva em: {caminho_saida}")
            logger.info(f"Planilha '{nome_aba}' salva em: {caminho_saida}")
    except Exception as e:
        print(f"⚠️ Erro ao processar planilhas do Excel: {e}")
        logger.error(f"Erro ao processar planilhas do Excel: {e}")
        return None

    print("🟢 [main_preparo_arquivos] Tudo pronto, retornando variáveis:")
    print("  TMP_DIR:", TMP_DIR)
    print("  PREPARADO:", PREPARADO)
    print("  CONCLUIDO:", CONCLUIDO)
    print("  Excel:", destino_excel)
    print("  DXF:", destino_dxf)
    print("  Template:", os.path.join(BASE_DIR, 'templates_doc', 'MD_DECOPA_PADRAO.docx'))
    logger.info("Preparo concluído com sucesso")

    return {
        "diretorio_final": TMP_DIR,
        "diretorio_preparado": PREPARADO,
        "diretorio_concluido": CONCLUIDO,
        "arquivo_excel_recebido": destino_excel,
        "arquivo_dxf_recebido": destino_dxf,
        "caminho_template": os.path.join(BASE_DIR, 'templates_doc', 'MD_DECOPA_PADRAO.docx')
    }
