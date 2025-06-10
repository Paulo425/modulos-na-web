import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import glob
import sys


# Solicitação clara e direta do diretório base pelo usuário
def selecionar_diretorio_base():
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()

    print("Selecione claramente o DIRETÓRIO BASE para salvar os arquivos.")
    diretorio_base = filedialog.askdirectory(title="Escolha o Diretório Base")

    root.destroy()
    return diretorio_base

# Função ajustada para criar diretórios necessários
def criar_diretorios(base, cidade):
    diretorio_cidade = os.path.join(base, cidade)

    if not os.path.exists(diretorio_cidade):
        for sub in ["RECEBIDO_CARLOS", "PREPARADO", "CONCLUIDO"]:
            os.makedirs(os.path.join(diretorio_cidade, sub), exist_ok=True)
    else:
        data_hoje = datetime.now().strftime("%d%b%y")
        diretorio_repescagem = os.path.join(diretorio_cidade, f"REPESCAGEM_{data_hoje}")
        for sub in ["RECEBIDO_CARLOS", "PREPARADO", "CONCLUIDO"]:
            os.makedirs(os.path.join(diretorio_repescagem, sub), exist_ok=True)
        diretorio_cidade = diretorio_repescagem

    return diretorio_cidade

# Função para seleção de arquivos (já correta)
def selecionar_arquivo(tipo):
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()

    if tipo.lower() == "dxf":
        tipos_arquivo = [("Arquivos DXF", "*.dxf")]
        mensagem = "INSIRA O ARQUIVO DXF ORIGINAL (.dxf)"
    elif tipo.lower() == "excel":
        tipos_arquivo = [("Arquivos Excel", "*.xlsx")]
        mensagem = "INSIRA O ARQUIVO EXCEL 'DADOS DO IMÓVEL' (.xlsx)"
    else:
        tipos_arquivo = [("Todos os arquivos", "*.*")]
        mensagem = f"INSIRA O ARQUIVO {tipo}"

    print(mensagem)
    caminho = filedialog.askopenfilename(title=mensagem, filetypes=tipos_arquivo)
    root.destroy()
    return caminho

# Função preparo das planilhas (já correta)
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

# Execução inicial revisada definitivamente
def main_preparo_arquivos(diretorio_saida, cidade, caminho_excel, caminho_dxf):
    print("🔵 Preparação inicial dos arquivos e diretórios 🔵")

    diretorio_base = diretorio_saida  # Agora usa o parâmetro recebido diretamente.

    # Agora utiliza diretamente a cidade passada como parâmetro (sem pedir input novamente)
    diretorio_final = criar_diretorios(diretorio_base, cidade)

    # Definição clara dos subdiretórios
    diretorio_recebido_carlos = os.path.join(diretorio_final, "RECEBIDO_CARLOS")
    diretorio_preparado = os.path.join(diretorio_final, "PREPARADO")
    diretorio_concluido = os.path.join(diretorio_final, "CONCLUIDO")

    print(f"📂 Estrutura criada com sucesso em: {diretorio_final}")

    # Agora usa diretamente os caminhos recebidos sem janela gráfica
    arquivo_excel_recebido = shutil.copy2(caminho_excel, diretorio_recebido_carlos)
    arquivo_dxf_recebido = shutil.copy2(caminho_dxf, diretorio_recebido_carlos)

    print(f"✅ Excel copiado para: {arquivo_excel_recebido}")
    print(f"✅ DXF copiado para: {arquivo_dxf_recebido}")

    # Preparação das planilhas
    preparar_planilhas(arquivo_excel_recebido, diretorio_preparado)
    # Ajusta dinamicamente para o PyInstaller:
    if hasattr(sys, '_MEIPASS'):
        diretorio_template = sys._MEIPASS
    else:
        diretorio_template = os.path.abspath(".")

    caminho_template = os.path.join(diretorio_template, "MD_DECOPA_PADRAO.docx")
    
    # Retorno claro e útil das variáveis
    variaveis_retorno = {
        "diretorio_final": diretorio_final,
        "cidade": cidade,
        "arquivo_excel_recebido": arquivo_excel_recebido,
        "arquivo_dxf_recebido": arquivo_dxf_recebido,
        "diretorio_preparado": diretorio_preparado,
        "diretorio_concluido": diretorio_concluido,
        "caminho_template": caminho_template
    }

    print("\n✅ Preparação concluída com sucesso! Variáveis disponíveis para uso posterior:")
    for chave, valor in variaveis_retorno.items():
        print(f"- {chave}: {valor}")

    return variaveis_retorno
