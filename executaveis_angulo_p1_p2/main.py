import os
import sys
import logging
from datetime import datetime
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
from poligonal_aberta import main_poligonal_aberta
from unir_poligonais import main_unir_poligonais

import shutil
import uuid
import subprocess


# ✅ 1. Caminho base
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


# ✅ 2. Pastas públicas
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

# ✅ 3. Pasta de logs
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(LOG_DIR, f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# Configura logging para arquivo e console (stream)
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Handler para arquivo
file_handler = logging.FileHandler(log_path, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Handler para console (StreamHandler)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)

# Formatação comum aos dois handlers
formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Adiciona handlers ao logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# ✅ 4. Habilita UTF-8 no console
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass  # Em alguns ambientes, reconfigure não está disponível

def main():
    if len(sys.argv) != 4:
        print("Uso: python main.py <cidade> <caminho_excel> <caminho_dxf>")
        sys.exit(1)

    cidade = sys.argv[1]
    uuid_str = str(uuid.uuid4())[:8]
    cidade_formatada = cidade.replace(" ", "_")
    caminho_excel = sys.argv[2]
    caminho_dxf = sys.argv[3]
    caminho_template = os.path.join(BASE_DIR, "templates_doc", "Memorial_modelo_padrao.docx")

    if not os.path.exists(caminho_template):
        print(f"Template '{caminho_template}' não encontrado.")
        sys.exit(1)

    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, BASE_DIR, uuid_str)

    # 🔷 Iniciando: Processamento Poligonal Aberta (corrigido)
    main_poligonal_aberta(
        uuid_str,
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"]
       
    )

    # 🔷 Iniciando: Processamento Poligonal Fechada
    main_poligonal_fechada(
        uuid_str,
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"],
        caminho_template
    )

    # 🔸 Unir poligonais
    main_unir_poligonais(
        uuid_str,
        variaveis["diretorio_concluido"],
        caminho_template
    )

    # Compactação final
    main_compactar_arquivos(
        variaveis["diretorio_concluido"],
        cidade_formatada,
        uuid_str
    )

    print("✅ [main.py] Compactação finalizada com sucesso!")

    # Copiar ZIPs para static/arquivos
    try:
        zips_copiados = 0
        pasta_origem = variaveis["diretorio_concluido"]
        pasta_destino = CAMINHO_PUBLICO
        os.makedirs(pasta_destino, exist_ok=True)

        for arquivo in os.listdir(pasta_origem):
            if arquivo.lower().endswith(".zip"):
                origem = os.path.join(pasta_origem, arquivo)
                destino = os.path.join(pasta_destino, arquivo)
                shutil.copy2(origem, destino)
                print(f"📦 ZIP copiado: {arquivo}")
                zips_copiados += 1

        if zips_copiados == 0:
            print("⚠️ Nenhum ZIP encontrado para copiar.")
    except Exception as e:
        print(f"❌ Erro ao copiar ZIPs: {e}")


    

if __name__ == "__main__":
    main()
