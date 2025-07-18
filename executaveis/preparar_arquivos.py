# preparar_arquivos.py

import os
import shutil
import logging
import pandas as pd
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

# ✅ Função para gerar planilhas abertas e fechadas como no AZIMUTE_AZ
def preparar_planilhas(arquivo_recebido, diretorio_preparado):
    def processar_planilha(df, coluna_codigo, identificador, diretorio_destino):
        if coluna_codigo not in df.columns:
            print(f"⚠️ Coluna '{coluna_codigo}' não encontrada.")
            return

        df_v = df[df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)][[coluna_codigo, "Confrontante"]]
        df_outros = df[~df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)]

        df_v.to_excel(os.path.join(diretorio_destino, f"FECHADA_{identificador}.xlsx"), index=False)
        df_outros.to_excel(os.path.join(diretorio_destino, f"ABERTA_{identificador}.xlsx"), index=False)
        print(f"✅ Planilhas processadas para: {identificador}")

    xls = pd.ExcelFile(arquivo_recebido)
    for sheet_name, sufixo in [("ETE", "ETE"), ("Confrontantes_Remanescente", "REM"),
                               ("Confrontantes_Servidao", "SER"), ("Confrontantes_Acesso", "ACE")]:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            identificador = f"{os.path.splitext(os.path.basename(arquivo_recebido))[0]}_{sufixo}"
            processar_planilha(df, "Código", identificador, diretorio_preparado)
        else:
            print(f"⚠️ Planilha '{sheet_name}' não encontrada.")

# ✅ Função principal usada pelo main.py
def main_preparo_arquivos(diretorio_base, cidade, caminho_excel, caminho_dxf):
    TMP_DIR = os.path.join(BASE_DIR, 'tmp')
    RECEBIDO = os.path.join(TMP_DIR, 'RECEBIDO')
    PREPARADO = os.path.join(TMP_DIR, 'PREPARADO')
    CONCLUIDO = diretorio_base  # já vem com UUID do main.py


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

    # ✅ Gera planilhas abertas/fechadas (ponto-chave que estava faltando)
    try:
        preparar_planilhas(destino_excel, PREPARADO)
    except Exception as e:
        print(f"⚠️ Erro ao preparar planilhas: {e}")
        logger.error(f"Erro ao preparar planilhas: {e}")

    # ✅ Também salva as planilhas completas como antes
    try:
        df = pd.read_excel(destino_excel, sheet_name=None)
        for nome_aba, tabela in df.items():
            nome_arquivo = f"{nome_aba}_PREPARADO.xlsx"
            caminho_saida = os.path.join(PREPARADO, nome_arquivo)
            tabela.to_excel(caminho_saida, index=False)
            print(f"✅ Planilha '{nome_aba}' salva em: {caminho_saida}")
            logger.info(f"Planilha '{nome_aba}' salva em: {caminho_saida}")
    except Exception as e:
        print(f"⚠️ Erro ao salvar planilhas completas: {e}")
        logger.error(f"Erro ao salvar planilhas completas: {e}")
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
