# executaveis_angulo_p1_p2/main.py

import os
import sys
import argparse
import logging
import json
from datetime import datetime

# ---------------------------------------------------------
# Garante import local e BASE_DIR
# ---------------------------------------------------------
EXEC_DIR = os.path.dirname(os.path.abspath(__file__))
if EXEC_DIR not in sys.path:
    sys.path.insert(0, EXEC_DIR)
BASE_DIR = os.path.abspath(os.path.join(EXEC_DIR, '..'))
os.environ.setdefault("BASE_DIR", BASE_DIR)

# ---------------------------------------------------------
# Pré-pega --id-execucao/--diretorio antes do exec_ctx
# (para o exec_ctx resolver ID_EXECUCAO e DIR_* na importação)
# ---------------------------------------------------------
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
            try:
                parent = os.path.dirname(dir_arg.rstrip(os.sep))
                cand = os.path.basename(parent)
                if cand:
                    os.environ["ID_EXECUCAO"] = cand
            except Exception:
                pass

_prefetch_id_from_cli()

# ---------------------------------------------------------
# Contexto único da execução
# ---------------------------------------------------------
from exec_ctx import ID_EXECUCAO, DIR_RUN, DIR_REC, DIR_PREP, DIR_CONC, setup_logger

# Pipeline ANGULO_P1_P2
from preparar_arquivos import preparar_arquivos
from poligonal_aberta import main_poligonal_aberta
from poligonal_fechada import main_poligonal_fechada
from unir_poligonais import main_unir_poligonais
from compactar_arquivos import main_compactar_arquivos

# Logger unificado (arquivo: CONCLUIDO/exec_<uuid>.log + stdout)
logger = setup_logger("pipeline")

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf, sentido_poligonal):
    """
    Orquestra o pipeline do ANGULO_P1_P2 usando ID_EXECUCAO único (do Flask).
    Etapas: PREPARO -> ABERTA -> FECHADA -> UNIR -> ZIP -> RUN.json
    """
    dir_conc = os.path.abspath(diretorio_saida) if diretorio_saida else DIR_CONC
    cidade_fmt = (cidade or "").replace(" ", "_")

    logger.info("🚀 Início ANGULO_P1_P2 | ID=%s", ID_EXECUCAO)
    logger.info("📁 Entradas: dir=%s | cidade=%s | excel=%s | dxf=%s | sentido=%s",
                dir_conc, cidade, caminho_excel, caminho_dxf, sentido_poligonal)

    # 1) Preparo (gera <uuid>_FECHADA_{TIPO}.xlsx em PREPARADO)
    vars = preparar_arquivos(cidade, caminho_excel, caminho_dxf, BASE_DIR, ID_EXECUCAO)
    if not isinstance(vars, dict) or not vars:
        logger.error("❌ preparar_arquivos falhou.")
        return 2

    dir_prep = vars.get("diretorio_preparado", DIR_PREP)
    dir_conc = vars.get("diretorio_concluido", DIR_CONC)
    xls_in   = vars.get("arquivo_excel_recebido")
    dxf_in   = vars.get("arquivo_dxf_recebido")
    tpl      = vars.get("caminho_template") or os.path.join(BASE_DIR, "templates_doc", "Memorial_modelo_padrao.docx")

    logger.info("✅ Preparo ok. PREPARADO=%s | CONCLUIDO=%s", dir_prep, dir_conc)

    # 2) Poligonal ABERTA
    logger.info("🔷 Processamento Poligonal ABERTA")
    main_poligonal_aberta(
        ID_EXECUCAO,
        xls_in,
        dxf_in,
        dir_prep,
        dir_conc
    )
    logger.info("✅ Poligonal ABERTA concluída.")

    # 3) Poligonal FECHADA
    logger.info("🔷 Processamento Poligonal FECHADA")
    main_poligonal_fechada(
        ID_EXECUCAO,
        xls_in,
        dxf_in,
        dir_prep,
        dir_conc,
        tpl,
        sentido_poligonal
    )
    logger.info("✅ Poligonal FECHADA concluída.")

    # 4) Unir poligonais
    logger.info("🔷 Unindo poligonais")
    main_unir_poligonais(
        ID_EXECUCAO,
        dir_conc,
        tpl
    )
    logger.info("✅ União das poligonais concluída.")

    # 5) Compactar (gera <uuid>_FECHADA_{TIPO}_{MATR}.zip em CONCLUIDO)
    logger.info("📦 Compactação: %s", dir_conc)
    main_compactar_arquivos(dir_conc, cidade_fmt, ID_E_
