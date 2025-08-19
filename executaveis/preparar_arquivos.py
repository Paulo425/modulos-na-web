# preparar_arquivos.py

import os
import shutil
import logging
import pandas as pd
from datetime import datetime

# Integração com o contexto único da execução
from exec_ctx import ID_EXECUCAO, DIR_RUN, DIR_REC, DIR_PREP, DIR_CONC

# ----------------------------------------------------------------------
# Diretórios e logger (mantidos no padrão do seu projeto)
# ----------------------------------------------------------------------
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

# ----------------------------------------------------------------------
# ✅ Função para gerar planilhas ABERTA/FECHADA (como no AZIMUTE_AZ)
# ----------------------------------------------------------------------
def preparar_planilhas(arquivo_recebido, diretorio_preparado):
    def processar_planilha(df, coluna_codigo, identificador, diretorio_destino):
        if coluna_codigo not in df.columns:
            print(f"⚠️ Coluna '{coluna_codigo}' não encontrada.")
            return

        # Linhas cujo código é V1, V2, ..., (ou v1, v2, ...)
        mask_vertices = df[coluna_codigo].astype(str).str.match(r'^[Vv][0-9]*$', na=False)
        df_v = df[mask_vertices][[coluna_codigo, "Confrontante"]] if "Confrontante" in df.columns else df[mask_vertices][[coluna_codigo]]
        df_outros = df[~mask_vertices]

        os.makedirs(diretorio_destino, exist_ok=True)
        df_v.to_excel(os.path.join(diretorio_destino, f"FECHADA_{identificador}.xlsx"), index=False)
        df_outros.to_excel(os.path.join(diretorio_destino, f"ABERTA_{identificador}.xlsx"), index=False)
        print(f"✅ Planilhas processadas para: {identificador}")

    xls = pd.ExcelFile(arquivo_recebido)
    for sheet_name, sufixo in [
        ("ETE", "ETE"),
        ("Confrontantes_Remanescente", "REM"),
        ("Confrontantes_Servidao", "SER"),
        ("Confrontantes_Acesso", "ACE"),
    ]:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            identificador = f"{os.path.splitext(os.path.basename(arquivo_recebido))[0]}_{sufixo}"
            processar_planilha(df, "Código", identificador, diretorio_preparado)
        else:
            print(f"⚠️ Planilha '{sheet_name}' não encontrada.")


# ----------------------------------------------------------------------
# ✅ Função principal usada pelo main.py (ASSINATURA MANTIDA)
# ----------------------------------------------------------------------
def main_preparo_arquivos(diretorio_saida, cidade, caminho_excel, caminho_dxf):
    """
    Prepara diretórios da execução, copia Excel/DXF para RECEBIDO,
    gera planilhas ABERTA/FECHADA no PREPARADO e salva todas as abas no PREPARADO.

    Retorna dict com:
      - diretorio_preparado
      - diretorio_concluido
      - arquivo_excel_recebido
      - arquivo_dxf_recebido
      - caminho_template
    """

    # ------------------------------------------------------------------
    # Diretórios desta execução (isolados pelo ID_EXECUCAO)
    # Mantemos nomes locais para retorno, mas garantimos consistência com exec_ctx
    # ------------------------------------------------------------------
    CONCLUIDO = os.path.abspath(diretorio_saida) if diretorio_saida else DIR_CONC
    # Se o diretório passado não corresponder ao ID da execução atual, força DIR_CONC
    expected_id = os.path.basename(os.path.dirname(CONCLUIDO))
    if expected_id != os.path.basename(DIR_RUN) or not CONCLUIDO.endswith("CONCLUIDO"):
        CONCLUIDO = DIR_CONC

    RUN_DIR   = DIR_RUN
    RECEBIDO  = DIR_REC
    PREPARADO = DIR_PREP

    for pasta in (RECEBIDO, PREPARADO, CONCLUIDO):
        os.makedirs(pasta, exist_ok=True)

    # ------------------------------------------------------------------
    # Validações mínimas de entrada
    # ------------------------------------------------------------------
    if not caminho_excel or not os.path.isfile(caminho_excel):
        msg = f"Arquivo Excel inválido ou inexistente: {caminho_excel}"
        print(f"❌ {msg}")
        logger.error(msg)
        return None

    if not caminho_dxf or not os.path.isfile(caminho_dxf):
        msg = f"Arquivo DXF inválido ou inexistente: {caminho_dxf}"
        print(f"❌ {msg}")
        logger.error(msg)
        return None

    # ------------------------------------------------------------------
    # Cópias para RECEBIDO
    # ------------------------------------------------------------------
    nome_excel = os.path.basename(caminho_excel)
    nome_dxf   = os.path.basename(caminho_dxf)

    destino_excel = os.path.join(RECEBIDO, nome_excel)
    destino_dxf   = os.path.join(RECEBIDO, nome_dxf)

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

    # ------------------------------------------------------------------
    # Gera planilhas ABERTA/FECHADA (ponto-chave)
    # ------------------------------------------------------------------
    try:
        preparar_planilhas(destino_excel, PREPARADO)
    except Exception as e:
        print(f"⚠️ Erro ao preparar planilhas: {e}")
        logger.error(f"Erro ao preparar planilhas: {e}")

    # ------------------------------------------------------------------
    # Salva todas as abas originais no PREPARADO
    # ------------------------------------------------------------------
    try:
        df_sheets = pd.read_excel(destino_excel, sheet_name=None)
        for nome_aba, tabela in df_sheets.items():
            nome_arquivo = f"{nome_aba}_PREPARADO.xlsx"
            caminho_saida = os.path.join(PREPARADO, nome_arquivo)
            tabela.to_excel(caminho_saida, index=False)
            print(f"✅ Planilha '{nome_aba}' salva em: {caminho_saida}")
            logger.info(f"Planilha '{nome_aba}' salva em: {caminho_saida}")
    except Exception as e:
        print(f"⚠️ Erro ao salvar planilhas completas: {e}")
        logger.error(f"Erro ao salvar planilhas completas: {e}")
        return None

    # ------------------------------------------------------------------
    # Logs de depuração (mantidos)
    # ------------------------------------------------------------------
    print("  ID_EXECUCAO:", ID_EXECUCAO)
    print("  RUN_DIR:", DIR_RUN)
    print("  PREPARADO:", DIR_PREP)
    print("  CONCLUIDO:", DIR_CONC)
    print("  Excel:", destino_excel)
    print("  DXF:", destino_dxf)

    caminho_template = os.path.join(BASE_DIR, 'templates_doc', 'MD_DECOPA_PADRAO.docx')
    print("  Template:", caminho_template)
    logger.info("Preparo concluído com sucesso")

    return {
        "diretorio_preparado": PREPARADO,
        "diretorio_concluido": CONCLUIDO,
        "arquivo_excel_recebido": destino_excel,
        "arquivo_dxf_recebido": destino_dxf,
        "caminho_template": caminho_template,
    }
