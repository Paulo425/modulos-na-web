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
if not any(isinstance(h, logging.FileHandler) and getattr(h, "baseFilename", None) == log_file for h in logger.handlers):
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    logger.addHandler(file_handler)

def preparar_planilhas(arquivo_recebido, diretorio_preparado, id_execucao):
    """
    Gera apenas as planilhas FECHADAS por tipo, no padrão:
      FECHADA_{<identificador>}_{TIPO}.xlsx
    Ex.: FECHADA_Dados_do_Imóvel_-_ETE_CHUI_-_Transcrição_32.681_ETE.xlsx
    """
    def processar_planilha(df, coluna_codigo, sufixo, diretorio_destino, identificador):
        if coluna_codigo not in df.columns:
            print(f"⚠️ Coluna '{coluna_codigo}' não encontrada.")
            logger.warning("Coluna '%s' não encontrada em %s", coluna_codigo, sufixo)
            return

        # Apenas vértices V1, V2, ... (poligonal FECHADA)
        mask_vertices = df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)
        if "Confrontante" in df.columns:
            df_v = df.loc[mask_vertices, [coluna_codigo, "Confrontante"]]
        else:
            df_v = df.loc[mask_vertices, [coluna_codigo]]  # tolera ausência de 'Confrontante'

        os.makedirs(diretorio_destino, exist_ok=True)

        # ⚠️ Padrão que as próximas etapas esperam: começa com "FECHADA_"
        nome_saida = f"FECHADA_{identificador}_{sufixo}.xlsx"
        caminho_saida = os.path.join(diretorio_destino, nome_saida)
        df_v.to_excel(caminho_saida, index=False)
        print(f"✅ Planilha FECHADA processada para: {sufixo} -> {caminho_saida}")
        logger.info("Planilha FECHADA gerada: %s", caminho_saida)

    xls = pd.ExcelFile(arquivo_recebido)
    for sheet_name, sufixo in [
        ("ETE", "ETE"),
        ("Confrontantes_Remanescente", "REM"),
        ("Confrontantes_Servidao", "SER"),
        ("Confrontantes_Acesso", "ACE"),
    ]:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            identificador = os.path.splitext(os.path.basename(arquivo_recebido))[0]
            processar_planilha(df, "Código", sufixo, diretorio_preparado, identificador)
        else:
            print(f"⚠️ Planilha '{sheet_name}' não encontrada.")
            logger.warning("Planilha '%s' não encontrada no Excel recebido.", sheet_name)

def preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir, id_execucao):
    """
    Prepara diretórios da execução, copia Excel/DXF para RECEBIDO e
    cria planilhas FECHADAS no PREPARADO. Retorna um dict com
    { arquivo_excel_recebido, arquivo_dxf_recebido, diretorio_preparado,
      diretorio_concluido, cidade_formatada } ou None em caso de erro.
    """
    try:
        cidade_formatada = (cidade or "").replace(" ", "_")

        TMP_DIR = os.path.join(base_dir, 'tmp', id_execucao)
        RECEBIDO = os.path.join(TMP_DIR, "RECEBIDO")
        PREPARADO = os.path.join(TMP_DIR, "PREPARADO")
        CONCLUIDO = os.path.join(TMP_DIR, "CONCLUIDO")

        for pasta in (RECEBIDO, PREPARADO, CONCLUIDO):
            os.makedirs(pasta, exist_ok=True)

        if not caminho_excel or not os.path.isfile(caminho_excel):
            msg = f"Arquivo Excel inválido: {caminho_excel}"
            print("❌", msg); logger.error(msg); return None
        if not caminho_dxf or not os.path.isfile(caminho_dxf):
            msg = f"Arquivo DXF inválido: {caminho_dxf}"
            print("❌", msg); logger.error(msg); return None

        nome_excel = os.path.basename(caminho_excel)
        nome_dxf   = os.path.basename(caminho_dxf)

        destino_excel = os.path.join(RECEBIDO, nome_excel)
        destino_dxf   = os.path.join(RECEBIDO, nome_dxf)

        shutil.copy2(caminho_excel, destino_excel)
        print(f"✅ Arquivo Excel copiado para: {destino_excel}")
        logger.info("Excel copiado: %s", destino_excel)

        shutil.copy2(caminho_dxf, destino_dxf)
        print(f"✅ Arquivo DXF copiado para: {destino_dxf}")
        logger.info("DXF copiado: %s", destino_dxf)

        # Gera FECHADAS no PREPARADO
        preparar_planilhas(destino_excel, PREPARADO, id_execucao)

        # Debug útil
        print("  ID_EXECUCAO:", id_execucao)
        print("  RECEBIDO   :", RECEBIDO)
        print("  PREPARADO  :", PREPARADO)
        print("  CONCLUIDO  :", CONCLUIDO)
        print("  Excel      :", destino_excel)
        print("  DXF        :", destino_dxf)

        return {
            "arquivo_excel_recebido": destino_excel,
            "arquivo_dxf_recebido": destino_dxf,
            "diretorio_preparado": PREPARADO,
            "diretorio_concluido": CONCLUIDO,
            "cidade_formatada": cidade_formatada,
            # (opcional) "caminho_template": os.path.join(BASE_DIR, "templates_doc", "Memorial_modelo_padrao.docx"),
        }

    except Exception as e:
        logger.exception("Erro ao preparar os arquivos: %s", e)
        print(f"❌ Erro ao preparar os arquivos: {e}")
        return None
