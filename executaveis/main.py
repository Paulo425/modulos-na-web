import argparse
import sys
import os
import time
import logging
import shutil
from datetime import datetime
from preparar_arquivos import main_preparo_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
import uuid
import json


# ✅ 1. Caminho base
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# 👉 Cada execução em diretórios próprios
#RUN_UUID = os.environ.get("RUN_UUID") or uuid.uuid4().hex[:8]
DIR_RUN  = os.path.join(BASE_DIR, 'tmp', RUN_UUID)
DIR_REC  = os.path.join(DIR_RUN, 'RECEBIDO')
DIR_PREP = os.path.join(DIR_RUN, 'PREPARADO')
DIR_CONC = os.path.join(DIR_RUN, 'CONCLUIDO')
for d in (DIR_REC, DIR_PREP, DIR_CONC):
    os.makedirs(d, exist_ok=True)

# log desta execução dentro do CONCLUIDO (onde o Flask procura)
log_path = os.path.join(DIR_CONC, f"exec_{RUN_UUID}.log")


# ✅ 2. Pastas públicas
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

# ✅ 3. Pasta de logs
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(LOG_DIR, f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# ✅ Logger raiz aponta para o log da execução + console
root = logging.getLogger()
root.setLevel(logging.DEBUG)
for h in list(root.handlers):
    root.removeHandler(h)

fh = logging.FileHandler(log_path, encoding='utf-8')
sh = logging.StreamHandler(sys.stdout)
fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
fh.setFormatter(fmt); sh.setFormatter(fmt)
root.addHandler(fh); root.addHandler(sh)


# ✅ 5. Habilita UTF-8 no console (com fallback para ambientes sem suporte)
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf):
    id_execucao = os.path.basename(diretorio_saida)  # ou 'diretorio' se for o nome da variável recebida

    print("🚀 [main.py] Início da execução principal")
    logging.info("🚀 Início da execução principal")

    print(f"📁 Variáveis de entrada: {diretorio_saida=}, {cidade=}, {caminho_excel=}, {caminho_dxf=}")
    logging.info(f"📁 Variáveis de entrada: {diretorio_saida=}, {cidade=}, {caminho_excel=}, {caminho_dxf=}")
    

    print("\n🔷 Iniciando: Preparo inicial dos arquivos")
    logging.info("🔷 Iniciando preparo inicial dos arquivos")

    variaveis = main_preparo_arquivos(diretorio_saida, cidade, caminho_excel, caminho_dxf)
    

    if not isinstance(variaveis, dict):
        print("❌ [main.py] ERRO: main_preparo_arquivos não retornou dicionário!")
        logging.error("❌ ERRO: main_preparo_arquivos não retornou dicionário!")
        return

    diretorio_preparado = variaveis["diretorio_preparado"]
    diretorio_concluido = variaveis["diretorio_concluido"]
    arquivo_excel_recebido = variaveis["arquivo_excel_recebido"]
    arquivo_dxf_recebido = variaveis["arquivo_dxf_recebido"]
    caminho_template = variaveis["caminho_template"]

    print("✅ [main.py] Preparo concluído. Variáveis carregadas.")
    logging.info("✅ Preparo concluído. Variáveis carregadas.")

    print("\n🔷 Processamento Poligonal Fechada")
    logging.info("🔷 Processamento Poligonal Fechada")

    main_poligonal_fechada(
        arquivo_excel_recebido,
        arquivo_dxf_recebido,
        diretorio_preparado,
        diretorio_concluido,
        caminho_template
    )

    print(f"\n📦 [main.py] Chamando compactação no diretório: {diretorio_concluido}")
    logging.info(f"📦 Chamando compactação no diretório: {diretorio_concluido}")

    main_compactar_arquivos(diretorio_concluido, cidade)



    print("✅ [main.py] Compactação finalizada com sucesso!")
    logging.info("✅ Compactação finalizada com sucesso!")

    
    try:
        zip_files = [f for f in os.listdir(diretorio_concluido) if f.lower().endswith('.zip')]
        with open(os.path.join(diretorio_concluido, "RUN.json"), "w", encoding="utf-8") as f:
            json.dump({"zip_files": zip_files}, f, ensure_ascii=False)
        logging.info(f"[RUN.json] registrado: {zip_files}")
    except Exception as e:
        logging.exception(f"Falha ao escrever RUN.json: {e}")


    print("✅ [main.py] Processo geral concluído com sucesso!")
    logging.info("✅ Processo geral concluído com sucesso!")
    print(f"📝 Log salvo em: static/logs/{os.path.basename(log_path)}")


if __name__ == "__main__":
    print("⚙️ [main.py] Script chamado diretamente via linha de comando")
    logging.info("⚙️ Script chamado diretamente via linha de comando")

    parser = argparse.ArgumentParser(description='Executar DECOPA diretamente com parâmetros.')
    parser.add_argument('--diretorio', help='Diretório onde salvar arquivos.')
    parser.add_argument('--cidade', help='Cidade do memorial.')
    parser.add_argument('--excel', help='Caminho do arquivo Excel.')
    parser.add_argument('--dxf', help='Caminho do arquivo DXF.')

    args = parser.parse_args()

    diretorio = args.diretorio
    cidade    = args.cidade
    excel     = args.excel
    dxf       = args.dxf
    


    # Ignoramos --diretorio aqui para padronizar por execução
    diretorio = DIR_CONC
    print(f"[DEBUG main.py] RUN_UUID: {RUN_UUID}")


    executar_programa(diretorio, cidade, excel, dxf)
