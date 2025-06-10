# preparar_arquivos.py

import os
import shutil
import pandas as pd

def main_preparo_arquivos(diretorio_base, cidade, caminho_excel, caminho_dxf):
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    TMP_DIR = os.path.join(BASE_DIR, 'tmp')
    RECEBIDO = os.path.join(TMP_DIR, 'RECEBIDO')
    PREPARADO = os.path.join(TMP_DIR, 'PREPARADO')
    CONCLUIDO = os.path.join(TMP_DIR, 'CONCLUIDO')

    # Criar subdiretórios se não existirem
    os.makedirs(RECEBIDO, exist_ok=True)
    os.makedirs(PREPARADO, exist_ok=True)
    os.makedirs(CONCLUIDO, exist_ok=True)

    # Copiar os arquivos recebidos para RECEBIDO
    nome_excel = os.path.basename(caminho_excel)
    nome_dxf = os.path.basename(caminho_dxf)

    destino_excel = os.path.join(RECEBIDO, nome_excel)
    destino_dxf = os.path.join(RECEBIDO, nome_dxf)

    shutil.copy(caminho_excel, destino_excel)
    shutil.copy(caminho_dxf, destino_dxf)

    print(f"✅ Excel copiado para: {destino_excel}")
    print(f"✅ DXF copiado para: {destino_dxf}")

    # Aqui você pode processar o Excel e gerar arquivos em PREPARADO
    # Exemplo: salvar uma planilha dividida em arquivos individuais
    try:
        df = pd.read_excel(destino_excel, sheet_name=None)
        for nome_aba, tabela in df.items():
            nome_arquivo = f"{nome_aba}_PREPARADO.xlsx"
            caminho_saida = os.path.join(PREPARADO, nome_arquivo)
            tabela.to_excel(caminho_saida, index=False)
            print(f"✅ Planilha '{nome_aba}' salva em: {caminho_saida}")
    except Exception as e:
        print(f"⚠️ Erro ao processar planilhas: {e}")

    # Retornar os caminhos úteis para as próximas fases
    return {
        "diretorio_final": TMP_DIR,
        "diretorio_preparado": PREPARADO,
        "diretorio_concluido": CONCLUIDO,
        "arquivo_excel_recebido": destino_excel,
        "arquivo_dxf_recebido": destino_dxf,
        "caminho_template": os.path.join(BASE_DIR, 'templates_doc', 'MD_DECOPA_PADRAO.docx')
    }
