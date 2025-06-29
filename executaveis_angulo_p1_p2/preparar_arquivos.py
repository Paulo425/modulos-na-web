import os
import shutil
import tempfile
import logging
import pandas as pd
import uuid
from datetime import datetime



# Diretórios e logger
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def preparar_planilhas(arquivo_recebido, diretorio_preparado):
    def processar_planilha(df, coluna_codigo, identificador, diretorio_destino, id_execucao):
        if coluna_codigo not in df.columns:
            mensagem = f"⚠️ Coluna '{coluna_codigo}' não encontrada na planilha '{identificador}'."
            print(mensagem)
            logger.warning(mensagem)
            return

        df_v = df[df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)][[coluna_codigo, "Confrontante"]]
        df_outros = df[~df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)]

        df_v.to_excel(os.path.join(diretorio_destino, f"{uuid_str}_FECHADA_{identificador}.xlsx"), index=False)
        df_outros.to_excel(os.path.join(diretorio_destino, f"{uuid_str}_ABERTA_{identificador}.xlsx"), index=False)

        logger.info(f"✅ Planilhas FECHADA e ABERTA geradas com UUID para identificador: {identificador}")
        print(f"✅ Planilhas FECHADA e ABERTA geradas com UUID para: {identificador}")

    xls = pd.ExcelFile(arquivo_recebido)
    for sheet_name, sufixo in [("ETE", "ETE"), ("Confrontantes_Remanescente", "REM"),
                               ("Confrontantes_Servidao", "SER"), ("Confrontantes_Acesso", "ACE")]:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            identificador = f"{os.path.splitext(os.path.basename(arquivo_recebido))[0]}_{sufixo}"
            processar_planilha(df, "Código", identificador, diretorio_preparado, id_execucao)
        else:
            mensagem = f"⚠️ Planilha '{sheet_name}' não encontrada no arquivo Excel."
            print(mensagem)
            logger.warning(mensagem)

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

        preparar_planilhas(destino_excel, PREPARADO)

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
