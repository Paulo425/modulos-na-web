# executaveis_azimute_az/main.py

import os
import sys
import argparse
import logging
import json
from datetime import datetime

# --- garantir import local de módulos irmãos (exec_ctx, preparar_arquivos, etc.) ---
EXEC_DIR = os.path.dirname(os.path.abspath(__file__))
if EXEC_DIR not in sys.path:
    sys.path.insert(0, EXEC_DIR)

# BASE_DIR (um nível acima da pasta executável do módulo)
BASE_DIR = os.path.abspath(os.path.join(EXEC_DIR, '..'))
os.environ.setdefault("BASE_DIR", BASE_DIR)

# -----------------------------------------------------------------------------
# Pré-leitura de argumentos para capturar --id-execucao/--diretorio ANTES do exec_ctx
# (permite ao exec_ctx resolver ID_EXECUCAO e DIR_* corretamente ao importar)
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
            # tenta extrair .../tmp/<ID>/CONCLUIDO
            try:
                parent = os.path.dirname(dir_arg.rstrip(os.sep))
                candidate = os.path.basename(parent)
                if candidate:
                    os.environ["ID_EXECUCAO"] = candidate
            except Exception:
                pass

_prefetch_id_from_cli()

# Agora podemos importar o contexto único da execução
from exec_ctx import ID_EXECUCAO, DIR_RUN, DIR_REC, DIR_PREP, DIR_CONC, setup_logger

# Módulos do pipeline deste módulo AZIMUTE_AZ
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

# Logger unificado (arquivo: CONCLUIDO/exec_<uuid>.log + stdout)
logger = setup_logger("pipeline")

# Pasta pública para eventual cópia (opcional)
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

# UTF-8 no console (se suportado)
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf, sentido_poligonal):
    """
    Orquestra o pipeline do AZIMUTE_AZ, usando o ID_EXECUCAO único vindo do Flask.
    """
    # Se não vier diretório, use o CONCLUIDO da execução corrente
    diretorio_concluido = os.path.abspath(diretorio_saida) if diretorio_saida else DIR_CONC
    cidade_formatada = (cidade or "").replace(" ", "_")

    logger.info("🚀 Início da execução AZIMUTE_AZ | ID=%s", ID_EXECUCAO)
    logger.info("📁 Entradas: diretorio=%s | cidade=%s | excel=%s | dxf=%s | sentido=%s",
                diretorio_concluido, cidade, caminho_excel, caminho_dxf, sentido_poligonal)

    # -------------------------------
    # 1) Preparo de arquivos
    # OBS: manter assinatura do seu preparar_arquivos original:
    # preparar_arquivos(cidade, caminho_excel, caminho_dxf, BASE_DIR, uuid_str)
    # Aqui passamos ID_EXECUCAO para manter consistência
    # -------------------------------
    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, BASE_DIR, ID_EXECUCAO)

    if not isinstance(variaveis, dict):
        logger.error("❌ preparar_arquivos não retornou dict. Abortando.")
        return 2

    diretorio_preparado     = variaveis.get("diretorio_preparado", DIR_PREP)
    diretorio_concluido     = variaveis.get("diretorio_concluido", DIR_CONC)
    arquivo_excel_recebido  = variaveis.get("arquivo_excel_recebido")
    arquivo_dxf_recebido    = variaveis.get("arquivo_dxf_recebido")
    caminho_template        = variaveis.get("caminho_template") or os.path.join(BASE_DIR, "templates_doc", "Memorial_modelo_padrao.docx")

    logger.info("✅ Preparo concluído. PREPARADO=%s | CONCLUIDO=%s", diretorio_preparado, diretorio_concluido)

    # -------------------------------
    # 2) Processar poligonal fechada (mantida a sua assinatura original)
    # main_poligonal_fechada(uuid, excel, dxf, pasta_preparado, pasta_concluido, template, sentido)
    # -------------------------------
    logger.info("🔷 Processamento Poligonal Fechada")
    main_poligonal_fechada(
        ID_EXECUCAO,
        arquivo_excel_recebido,
        arquivo_dxf_recebido,
        diretorio_preparado,
        diretorio_concluido,
        caminho_template,
        sentido_poligonal
    )
    logger.info("✅ Processamento da poligonal fechada concluído.")

    # -------------------------------
    # 3) Compactação
    # Assinatura que você usa: main_compactar_arquivos(dir_concluido, cidade_formatada, uuid)
    # -------------------------------
    logger.info("📦 Compactação: %s", diretorio_concluido)
    main_compactar_arquivos(diretorio_concluido, cidade_formatada, ID_EXECUCAO)
    logger.info("✅ Compactação concluída.")

    # -------------------------------
    # 4) RUN.json (manifesto) — redundância segura
    # -------------------------------
    try:
        zip_files = [f for f in os.listdir(diretorio_concluido) if f.lower().endswith('.zip')]
        run_json_path = os.path.join(diretorio_concluido, "RUN.json")
        with open(run_json_path, "w", encoding="utf-8") as f:
            json.dump({"zip_files": zip_files, "id_execucao": ID_EXECUCAO}, f, ensure_ascii=False)
        logger.info("[RUN.json] registrado: %s", zip_files)
    except Exception as e:
        logger.exception("Falha ao escrever RUN.json: %s", e)

    logger.info("✅ Processo geral concluído com sucesso!")
    return 0


def main():
    parser = argparse.ArgumentParser(description='Executar AZIMUTE_AZ com contexto de execução único (UUID).')
    parser.add_argument('--diretorio', help='Diretório CONCLUIDO desta execução. Padrão: DIR_CONC do exec_ctx.')
    parser.add_argument('--cidade', help='Cidade do memorial.')
    parser.add_argument('--excel', help='Caminho do arquivo Excel.')
    parser.add_argument('--dxf', help='Caminho do arquivo DXF.')
    parser.add_argument('--sentido', choices=['horario', 'anti_horario'], default='horario', help='Sentido da poligonal.')
    parser.add_argument('--id-execucao', help='ID único da execução (propagado pela rota Flask).')

    args = parser.parse_args()

    # Compatibilidade: se passou --id-execucao, reforça no ambiente (exec_ctx usa isso)
    if args.id_execucao:
        os.environ["ID_EXECUCAO"] = args.id_execucao

    # Defaults consistentes com exec_ctx
    diretorio = args.diretorio or DIR_CONC
    cidade    = args.cidade
    excel     = args.excel
    dxf       = args.dxf
    sentido   = args.sentido

    # Validações mínimas
    missing = []
    if not cidade: missing.append("--cidade")
    if not excel:  missing.append("--excel")
    if not dxf:    missing.append("--dxf")
    if missing:
        print("Uso incorreto. Faltando:", ", ".join(missing))
        parser.print_help()
        return 2

    rc = executar_programa(diretorio, cidade, excel, dxf, sentido)
    sys.exit(rc)


if __name__ == "__main__":
    main()
