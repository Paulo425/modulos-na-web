import os
import glob
import zipfile
import re
import logging
from datetime import datetime
from shutil import copyfile

# Diretórios e setup de log
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# Configurar logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def montar_pacote_zip(diretorio, cidade_formatada):
    tipos = ["ETE", "REM", "SER", "ACE"]
    for tipo in tipos:
        padrao_dxf = os.path.join(diretorio, f"{tipo}_Memorial_*.dxf")
        padrao_docx = os.path.join(diretorio, f"{tipo}_Memorial_MAT_*.docx")
        padrao_excel = os.path.join(diretorio, f"{tipo}_Memorial_*.xlsx")

        arquivo_dxf = glob.glob(padrao_dxf)
        arquivo_docx = glob.glob(padrao_docx)
        arquivo_excel = glob.glob(padrao_excel)

        if arquivo_dxf and arquivo_docx and arquivo_excel:
            # Extrai parte significativa do nome (por exemplo: "49_Transcrição 43.192")
            base_nome = os.path.splitext(os.path.basename(arquivo_dxf[0]))[0]
            partes = base_nome.split("_", 1)
            sufixo_identificador = partes[1] if len(partes) > 1 else partes[0]

            # ✅ Constrói o nome do ZIP conforme seu padrão
            nome_zip = f"{cidade_formatada}_{tipo}_{sufixo_identificador}.zip"
            caminho_zip = os.path.join(diretorio, nome_zip)

            with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                zipf.write(arquivo_dxf[0], os.path.basename(arquivo_dxf[0]))
                zipf.write(arquivo_docx[0], os.path.basename(arquivo_docx[0]))
                zipf.write(arquivo_excel[0], os.path.basename(arquivo_excel[0]))

            print(f"✅ Arquivos do tipo {tipo} compactados com sucesso!")
            print(f"🗜️ ZIP salvo em: {caminho_zip}")
        else:
            print(f"⚠️ Arquivos incompletos ou não encontrados para o tipo {tipo}")

def main_compactar_arquivos(diretorio_concluido, cidade_formatada):
    montar_pacote_zip(diretorio_concluido, cidade_formatada)
