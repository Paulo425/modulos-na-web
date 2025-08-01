import os
import shutil
import tempfile
import logging
import pandas as pd
import uuid
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

def preparar_planilhas(arquivo_recebido, diretorio_preparado, uuid_str):
    def processar_planilha(df, coluna_codigo, sufixo, diretorio_destino, uuid_str, identificador):
        if coluna_codigo not in df.columns:
            print(f"⚠️ Coluna '{coluna_codigo}' não encontrada.")
            return

        # Apenas os vértices V1, V2... são relevantes (Poligonal FECHADA)
        df_v = df[df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)][[coluna_codigo, "Confrontante"]]

        # Salva apenas a FECHADA
        df_v.to_excel(os.path.join(diretorio_destino, f"{uuid_str}_FECHADA_{sufixo}.xlsx"), index=False)

        print(f"✅ Planilha FECHADA processada para: {sufixo}")

        print(f"✅ Planilhas processadas para: {identificador}")

    xls = pd.ExcelFile(arquivo_recebido)
    for sheet_name, sufixo in [("ETE", "ETE"), ("Confrontantes_Remanescente", "REM"),
                               ("Confrontantes_Servidao", "SER"), ("Confrontantes_Acesso", "ACE")]:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            identificador = os.path.splitext(os.path.basename(arquivo_recebido))[0]
            processar_planilha(df, "Código", sufixo, diretorio_preparado, uuid_str, identificador)
        else:
            print(f"⚠️ Planilha '{sheet_name}' não encontrada.")

def preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir, id_execucao):
    try:
        cidade_formatada = cidade.replace(" ", "_")

        TMP_DIR = os.path.join(base_dir, 'tmp', id_execucao)
        os.makedirs(TMP_DIR, exist_ok=True)

        RECEBIDO = os.path.join(TMP_DIR, "RECEBIDO")
        PREPARADO = os.path.join(TMP_DIR, "PREPARADO")
        CONCLUIDO = os.path.join(TMP_DIR, "CONCLUIDO")

        for pasta in [RECEBIDO, PREPARADO, CONCLUIDO]:
            os.makedirs(pasta, exist_ok=True)

     

        nome_excel = os.path.basename(caminho_excel)
        nome_dxf = os.path.basename(caminho_dxf)

        destino_excel = os.path.join(RECEBIDO, nome_excel)
        destino_dxf = os.path.join(RECEBIDO, nome_dxf)

        shutil.copy(caminho_excel, destino_excel)
        print(f"✅ Arquivo Excel copiado para: {destino_excel}")
        shutil.copy(caminho_dxf, destino_dxf)
        print(f"✅ Arquivo DXF copiado para: {destino_dxf}")

        preparar_planilhas(destino_excel, PREPARADO, id_execucao)


        return {
            "arquivo_excel_recebido": destino_excel,
            "arquivo_dxf_recebido": destino_dxf,
            "diretorio_preparado": PREPARADO,
            "diretorio_concluido": CONCLUIDO,
            "cidade_formatada": cidade_formatada
        }

    except Exception as e:
        logger.error(f"Erro ao preparar os arquivos: {e}")
        print(f"❌ Erro ao preparar os arquivos: {e}")
        return {}
