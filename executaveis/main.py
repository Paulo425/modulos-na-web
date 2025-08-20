import os
import sys
import argparse
import logging
import shutil
from datetime import datetime
import json
import time

EXEC_DIR = os.path.dirname(os.path.abspath(__file__))
if EXEC_DIR not in sys.path:
    sys.path.insert(0, EXEC_DIR)
# ----------------------------------------------------------------------
# Base do projeto (um nível acima de executaveis) e repasse ao exec_ctx
# ----------------------------------------------------------------------
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.environ.setdefault("BASE_DIR", BASE_DIR)

# -----------------------------------------------------------------------------
# Pré-leitura de argumentos para capturar --id-execucao/--diretorio ANTES do exec_ctx
# -----------------------------------------------------------------------------
def _prefetch_id_from_cli():
    argv = sys.argv[1:]
    id_arg = None
    dir_arg = None
    for i, a in enumerate(argv):
        if a.startswith("--id-execucao="):
            id_arg = a.split("=", 1)[1].strip()
        elif a == "--id-execucao" and i + 1 < len(argv):
            id_arg = argv[i + 1].strip()
        elif a.startswith("--diretorio="):
            dir_arg = a.split("=", 1)[1].strip()
        elif a == "--diretorio" and i + 1 < len(argv):
            dir_arg = argv[i + 1].strip()

    if not os.environ.get("ID_EXECUCAO"):
        if id_arg:
            os.environ["ID_EXECUCAO"] = id_arg
        elif dir_arg:
            # tenta extrair padrão .../tmp/<ID>/CONCLUIDO
            try:
                parent = os.path.dirname(dir_arg.rstrip(os.sep))
                candidate = os.path.basename(parent)
                if candidate:
                    os.environ["ID_EXECUCAO"] = candidate
            except Exception:
                pass

_prefetch_id_from_cli()

# ----------------------------------------------------------------------
# Agora podemos importar o exec_ctx (usa BASE_DIR e ID_EXECUCAO acima)
# ----------------------------------------------------------------------
from exec_ctx import ID_EXECUCAO, DIR_RUN, DIR_REC, DIR_PREP, DIR_CONC, setup_logger
from preparar_arquivos import main_preparo_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

# ----------------------------------------------------------------------
# Pastas públicas e logger adicional em static/logs (mantido do seu padrão)
# ----------------------------------------------------------------------
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(LOG_DIR, f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

root = logging.getLogger()
root.setLevel(logging.DEBUG)
for h in list(root.handlers):
    root.removeHandler(h)

fh = logging.FileHandler(log_path, encoding='utf-8')
sh = logging.StreamHandler(sys.stdout)
fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
fh.setFormatter(fmt)
sh.setFormatter(fmt)
root.addHandler(fh)
root.addHandler(sh)

# Logger do exec_ctx (grava também em CONCLUIDO/exec_<ID>.log)
pipeline_logger = setup_logger("pipeline")

# UTF-8 no console (fallback se não suportar)
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf, sentido_poligonal='horario'):

    """
    Mantido com a mesma assinatura e variáveis de uso.
    """
    # Se nenhum diretório foi informado, usa o CONCLUIDO da execução corrente
    if not diretorio_saida:
        diretorio_saida = DIR_CONC

    id_execucao = ID_EXECUCAO  # fonte única da verdade

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

    diretorio_preparado     = variaveis["diretorio_preparado"]
    diretorio_concluido     = variaveis["diretorio_concluido"]
    arquivo_excel_recebido  = variaveis["arquivo_excel_recebido"]
    arquivo_dxf_recebido    = variaveis["arquivo_dxf_recebido"]
    caminho_template        = variaveis["caminho_template"]

    print("✅ [main.py] Preparo concluído. Variáveis carregadas.")
    logging.info("✅ Preparo concluído. Variáveis carregadas.")

    print("\n🔷 Processamento Poligonal Fechada")
    logging.info("🔷 Processamento Poligonal Fechada")

    main_poligonal_fechada(
        arquivo_excel_recebido,
        arquivo_dxf_recebido,
        diretorio_preparado,
        diretorio_concluido,
        caminho_template,
        sentido_poligonal
    )


    print(f"\n📦 [main.py] Chamando compactação no diretório: {diretorio_concluido}")
    logging.info(f"📦 Chamando compactação no diretório: {diretorio_concluido}")

    main_compactar_arquivos(diretorio_concluido, cidade)

    print("✅ [main.py] Compactação finalizada com sucesso!")
    logging.info("✅ Compactação finalizada com sucesso!")

    # RUN.json auxiliar
    try:
        zip_files = [f for f in os.listdir(diretorio_concluido) if f.lower().endswith('.zip')]
        with open(os.path.join(diretorio_concluido, "RUN.json"), "w", encoding="utf-8") as f:
            json.dump({"zip_files": zip_files, "id_execucao": id_execucao}, f, ensure_ascii=False)
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
    parser.add_argument('--id-execucao', help='ID único da execução (propagado pela rota Flask).')
    parser.add_argument('--sentido-poligonal', help='Sentido da poligonal (horario/antihorario).')

    args = parser.parse_args()

    diretorio = args.diretorio or DIR_CONC   # padrão consistente com exec_ctx
    cidade    = args.cidade
    excel     = args.excel
    dxf       = args.dxf
    sentido_poligonal = args.sentido_poligonal or 'horario'
    logger.info(f"Sentido poligonal recebido no main.py: {sentido_poligonal}")

    # Se passou --id-execucao aqui, garante no ambiente (compatibilidade)
    if args.id_execucao:
        os.environ["ID_EXECUCAO"] = args.id_execucao

    executar_programa(diretorio, cidade, excel, dxf, sentido_poligonal)

