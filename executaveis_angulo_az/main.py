# executaveis_angulo_az/main.py

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

# Pipeline ANGULO_AZ
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

# Logger unificado (arquivo: CONCLUIDO/exec_<uuid>.log + stdout)
logger = setup_logger("pipeline")

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf, sentido_poligonal):
    """
    Orquestra o pipeline do ANGULO_AZ usando ID_EXECUCAO único (do Flask).
    """
    dir_conc = os.path.abspath(diretorio_saida) if diretorio_saida else DIR_CONC
    cidade_fmt = (cidade or "").replace(" ", "_")

    logger.info("🚀 Início ANGULO_AZ | ID=%s", ID_EXECUCAO)
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

    # 2) Poligonal fechada (gera <uuid>_FECHADA_{TIPO}_{MATR}.xlsx/.docx/.dxf em CONCLUIDO)
    logger.info("🔷 Processamento Poligonal Fechada")
    main_poligonal_fechada(
        ID_EXECUCAO,
        xls_in,
        dxf_in,
        dir_prep,
        dir_conc,
        tpl,
        sentido_poligonal
    )
    logger.info("✅ Poligonal fechada concluída.")

    # 3) Compactar (gera <uuid>_FECHADA_{TIPO}_{MATR}.zip em CONCLUIDO)
    logger.info("📦 Compactação: %s", dir_conc)
    main_compactar_arquivos(dir_conc, cidade_fmt, ID_EXECUCAO)
    logger.info("✅ Compactação concluída.")

    # 4) RUN.json (redundância segura)
    try:
        created = [f for f in os.listdir(dir_conc) if f.lower().endswith(".zip")]
        run_json = os.path.join(dir_conc, "RUN.json")
        with open(run_json, "w", encoding="utf-8") as f:
            json.dump({"zip_files": created, "id_execucao": ID_EXECUCAO}, f, ensure_ascii=False)
        logger.info("[RUN.json] %s", created)
    except Exception as e:
        logger.exception("Falha RUN.json: %s", e)

    logger.info("✅ Processo geral concluído com sucesso!")
    return 0


def _parse_args():
    """
    Suporta:
    - Modo novo (nomeado): --id-execucao --diretorio --cidade --excel --dxf --sentido
    - Modo legado (posicional): main.py <cidade> <excel> <dxf> [sentido]
    """
    # Se NÃO há flags nomeadas e há 3-4 posicionais -> legado
    argv = sys.argv[1:]
    has_flags = any(a.startswith("--") for a in argv)
    if not has_flags and len(argv) in (3, 4):
        cidade, excel, dxf = argv[0], argv[1], argv[2]
        sentido = argv[3] if len(argv) == 4 else "horario"
        return argparse.Namespace(
            diretorio=DIR_CONC, cidade=cidade, excel=excel, dxf=dxf, sentido=sentido, id_execucao=os.environ.get("ID_EXECUCAO")
        )

    parser = argparse.ArgumentParser(description="Executar ANGULO_AZ com contexto de execução único.")
    parser.add_argument('--diretorio', help='Diretório CONCLUIDO (padrão: DIR_CONC do exec_ctx).')
    parser.add_argument('--cidade', required=True, help='Cidade do memorial.')
    parser.add_argument('--excel', required=True, help='Caminho do arquivo Excel.')
    parser.add_argument('--dxf', required=True, help='Caminho do arquivo DXF.')
    parser.add_argument('--sentido', choices=['horario', 'anti_horario'], default='horario', help='Sentido da poligonal.')
    parser.add_argument('--id-execucao', help='ID único da execução (propagado pelo Flask).')
    return parser.parse_args()


def main():
    args = _parse_args()

    # Compat: se passou --id-execucao, reforça no env (exec_ctx usa)
    if getattr(args, "id_execucao", None):
        os.environ["ID_EXECUCAO"] = args.id_execucao

    diretorio = args.diretorio or DIR_CONC
    cidade    = args.cidade
    excel     = args.excel
    dxf       = args.dxf
    sentido   = args.sentido

    # Validação mínima
    missing = []
    if not cidade: missing.append("--cidade")
    if not excel:  missing.append("--excel")
    if not dxf:    missing.append("--dxf")
    if missing:
        print("Uso incorreto. Faltando:", ", ".join(missing))
        return 2

    rc = executar_programa(diretorio, cidade, excel, dxf, sentido)
    sys.exit(rc)


if __name__ == "__main__":
    main()
