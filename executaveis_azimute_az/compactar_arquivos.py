import os
import glob
import zipfile
import re
import logging
import json
import shutil
from datetime import datetime

# Configuração inicial do diretório e log
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not any(isinstance(h, logging.FileHandler) and getattr(h, "baseFilename", None) == log_file for h in logger.handlers):
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    logger.addHandler(file_handler)


def montar_pacote_zip(diretorio, cidade_formatada, uuid_str):
    """
    Procura trios coerentes {TIPO}_{MATRICULA}.(xlsx/docx/dxf) e gera {CIDADE}_{TIPO}_{MATRICULA}.zip.
    Copia para static/arquivos com prefixo do UUID: {uuid}_{CIDADE}_{TIPO}_{MATRICULA}.zip
    e grava RUN.json em CONCLUIDO.
    """
    logger.info(f"Iniciando montagem dos pacotes ZIP em {diretorio}")
    try:
        lista = os.listdir(diretorio)
        logger.info(f"[DEBUG AZIMUTE_AZ] Arquivos no diretório antes do ZIP: {lista}")
    except Exception as e:
        logger.error(f"Falha ao listar diretório {diretorio}: {e}")
        return

    created_zips = []
    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        padrao_dxf = os.path.join(diretorio, f"*{tipo}*.dxf")
        padrao_docx = os.path.join(diretorio, f"*{tipo}*.docx")
        padrao_excel = os.path.join(diretorio, f"*{tipo}*.xlsx")

        arquivos_dxf = glob.glob(padrao_dxf)
        arquivos_docx = glob.glob(padrao_docx)
        arquivos_excel = glob.glob(padrao_excel)

        logger.info(f"[{tipo}] DXF={len(arquivos_dxf)} DOCX={len(arquivos_docx)} XLSX={len(arquivos_excel)}")

        # Extrai matrículas observadas nos nomes de arquivo
        matriculas = set()
        for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
            nome = os.path.basename(arq)
            m = re.search(rf"{tipo}.*?(\d{{2,6}}[.,_]?\d{{0,3}})", nome, re.IGNORECASE)
            if m:
                mat = m.group(1).replace(",", ".").replace(" ", "")
                if "." not in mat and len(mat) > 3:
                    mat = f"{mat[:-3]}.{mat[-3:]}"
                matriculas.add(mat)

        for matricula in matriculas:
            arq_dxf = [a for a in arquivos_dxf if matricula in a]
            arq_docx = [a for a in arquivos_docx if matricula in a]
            arq_excel = [a for a in arquivos_excel if matricula in a]

            if arq_dxf and arq_docx and arq_excel:
                cidade_sanit = (cidade_formatada or "CIDADE").replace(" ", "_")
                zip_name = f"{cidade_sanit}_{tipo}_{matricula}.zip"
                zip_path = os.path.join(diretorio, zip_name)

                try:
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        zf.write(arq_docx[0], os.path.basename(arq_docx[0]))
                        zf.write(arq_dxf[0],  os.path.basename(arq_dxf[0]))
                        zf.write(arq_excel[0], os.path.basename(arq_excel[0]))
                    logger.info(f"✅ ZIP criado: {zip_path}")
                    created_zips.append(zip_name)

                    # cópia pública com UUID no nome (anti-sobreposição)
                    try:
                        static_dir = os.path.join(BASE_DIR, 'static', 'arquivos')
                        os.makedirs(static_dir, exist_ok=True)
                        public_name = f"{uuid_str}_{zip_name}"
                        tmp_pub = os.path.join(static_dir, public_name + ".tmp")
                        shutil.copy2(zip_path, tmp_pub)
                        os.replace(tmp_pub, os.path.join(static_dir, public_name))
                        logger.info(f"🪣 ZIP também copiado (público): {public_name}")
                    except Exception as e_copy:
                        logger.warning(f"Falha ao copiar ZIP para público: {e_copy}")

                except Exception as e:
                    logger.exception(f"Erro ao criar ZIP {zip_path}: {e}")
            else:
                logger.info(f"[{tipo}/{matricula}] Arquivos insuficientes: "
                            f"DXF={len(arq_dxf)} DOCX={len(arq_docx)} XLSX={len(arq_excel)}")

    # Manifesto com os zips desta execução no CONCLUIDO
    try:
        run_json = os.path.join(diretorio, "RUN.json")
        with open(run_json, "w", encoding="utf-8") as f:
            json.dump({"zip_files": created_zips, "id_execucao": uuid_str}, f, ensure_ascii=False)
        logger.info(f"[RUN] Manifesto salvo: {run_json} | zip_files={created_zips}")
    except Exception as e:
        logger.warning(f"[RUN] Falha ao salvar RUN.json: {e}")

    logger.info("Compactação finalizada")


def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    logger.info("Compactação finalizada")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Compacta arquivos gerados (AZIMUTE_AZ).")
    parser.add_argument('--diretorio', required=True, help="Diretório dos arquivos concluídos.")
    parser.add_argument('--cidade', required=True, help="Nome formatado da cidade.")
    parser.add_argument('--uuid', required=True, help="UUID da execução atual.")
    args = parser.parse_args()

    try:
        logger.info(f"[DEBUG] Conteúdo de {args.diretorio}: {os.listdir(args.diretorio)}")
    except Exception as e:
        logger.warning(f"Não foi possível listar {args.diretorio}: {e}")

    main_compactar_arquivos(args.diretorio, args.cidade, args.uuid)
