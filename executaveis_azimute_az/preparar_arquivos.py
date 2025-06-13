import os
import shutil
import tempfile
import logging
import pandas as pd
# Configura o logger para funcionar no ambiente Render
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

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

def preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir):
    try:
        # Criar diretórios temporários (sem cidade no nome)
        diretorio_base = tempfile.mkdtemp()
        diretorio_preparado = os.path.join(diretorio_base, "PREPARADO")
        diretorio_concluido = os.path.join(diretorio_base, "CONCLUIDO")
        os.makedirs(diretorio_preparado, exist_ok=True)
        os.makedirs(diretorio_concluido, exist_ok=True)

        # Copiar arquivos para a pasta base
        arquivo_excel_recebido = os.path.join(diretorio_base, os.path.basename(caminho_excel))
        arquivo_dxf_recebido = os.path.join(diretorio_base, os.path.basename(caminho_dxf))
        shutil.copy(caminho_excel, arquivo_excel_recebido)
        shutil.copy(caminho_dxf, arquivo_dxf_recebido)

        print(f"✅ Arquivo Excel copiado para: {arquivo_excel_recebido}")
        print(f"✅ Arquivo DXF copiado para: {arquivo_dxf_recebido}")

        # Processar planilhas (gera arquivos na pasta PREPARADO)
        preparar_planilhas(arquivo_excel_recebido, diretorio_preparado)

        return {
            "arquivo_excel_recebido": arquivo_excel_recebido,
            "arquivo_dxf_recebido": arquivo_dxf_recebido,
            "diretorio_base": diretorio_base,
            "diretorio_preparado": diretorio_preparado,
            "diretorio_concluido": diretorio_concluido,
            "cidade_formatada": cidade_formatada
        }

    except Exception as e:
        logger.error(f"Erro ao preparar os arquivos: {e}")
        print(f"❌ Erro ao preparar os arquivos: {e}")
        return {}
