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

# ✅ 1. Caminho base
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# ✅ 2. Pastas públicas
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

# ✅ 3. Pasta de logs
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(LOG_DIR, f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# ✅ 4. Configura logger
logging.basicConfig(
    filename=log_path,
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
)

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

    print("\n📤 Copiando arquivos finais para a pasta pública")
    logging.info("📤 Copiando arquivos finais para a pasta pública")

    # Cria pasta pública se necessário
    os.makedirs(CAMINHO_PUBLICO, exist_ok=True)
    zip_download = None
    for fname in os.listdir(diretorio_concluido):
        origem = os.path.join(diretorio_concluido, fname)
        if os.path.isfile(origem):
            nome_com_uuid = f"{id_execucao}_{fname}"
            destino = os.path.join(BASE_DIR, 'static', 'arquivos', nome_com_uuid)
            try:
                shutil.copy2(origem, destino)
                print(f"🗂️ Arquivo copiado: {destino}")
                logging.info(f"🗂️ Arquivo copiado: {destino}")
                if fname.endswith(".zip"):
                    zip_download = nome_com_uuid
            except Exception as e:
                print(f"❌ Falha ao copiar {fname}: {e}")
                logging.error(f"❌ Erro ao copiar {fname}: {e}")


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
    cidade = args.cidade
    excel = args.excel
    dxf = args.dxf

    if not diretorio or 'C:\\' in diretorio or 'OneDrive' in diretorio:
        id_execucao = str(uuid.uuid4())[:8]
        diretorio= os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO', id_execucao)

    executar_programa(diretorio, cidade, excel, dxf)
