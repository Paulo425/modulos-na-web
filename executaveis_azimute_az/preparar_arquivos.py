import os
import shutil
import pandas as pd
from datetime import datetime

def preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir):
    diretorio_final = os.path.join(base_dir, cidade)
    if not os.path.exists(diretorio_final):
        for sub in ["RECEBIDO_CARLOS", "PREPARADO", "CONCLUIDO"]:
            os.makedirs(os.path.join(diretorio_final, sub), exist_ok=True)
    else:
        data_hoje = datetime.now().strftime("%d%b%y")
        diretorio_final = os.path.join(diretorio_final, f"REPESCAGEM_{data_hoje}")
        for sub in ["RECEBIDO_CARLOS", "PREPARADO", "CONCLUIDO"]:
            os.makedirs(os.path.join(diretorio_final, sub), exist_ok=True)

    diretorio_recebido_carlos = os.path.join(diretorio_final, "RECEBIDO_CARLOS")
    diretorio_preparado = os.path.join(diretorio_final, "PREPARADO")
    diretorio_concluido = os.path.join(diretorio_final, "CONCLUIDO")

    # Copiar os arquivos recebidos
    excel_dest = shutil.copy2(caminho_excel, diretorio_recebido_carlos)
    dxf_dest = shutil.copy2(caminho_dxf, diretorio_recebido_carlos)

    return {
        "diretorio_final": diretorio_final,
        "diretorio_recebido_carlos": diretorio_recebido_carlos,
        "diretorio_preparado": diretorio_preparado,
        "diretorio_concluido": diretorio_concluido,
        "arquivo_excel_recebido": excel_dest,
        "arquivo_dxf_recebido": dxf_dest
    }
