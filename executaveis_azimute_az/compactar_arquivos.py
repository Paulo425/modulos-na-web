import os
import glob
import zipfile
import logging
from datetime import datetime

# Configuração inicial do diretório e log
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

log_file = os.path.join(LOG_DIR, f"zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
logger.addHandler(file_handler)

def montar_pacote_zip(diretorio, cidade_formatada, uuid_str):
    """
    Procura trios coerentes {uuid}_FECHADA_{TIPO}_{MATRICULA}.(xlsx,docx,dxf)
    e gera {uuid}_FECHADA_{TIPO}_{MATRICULA}.zip
    """
    logger.info(f"Iniciando montagem dos pacotes ZIP em {diretorio}")
    try:
        logger.info(f"[DEBUG ANGULO_AZ] Arquivos no diretório antes do ZIP: {os.listdir(diretorio)}")
    except Exception as e:
        logger.error(f"Falha ao listar diretório {diretorio}: {e}")
        return

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        padrao_excel = os.path.join(diretorio, f"{uuid_str}_FECHADA_{tipo}_*.xlsx")
        excels = sorted(glob.glob(padrao_excel), key=os.path.getmtime, reverse=True)
        logger.info(f"[{tipo}] XLSX encontrados: {len(excels)} no padrão {padrao_excel}")

        for excel_path in excels:
            base = os.path.basename(excel_path)
            prefixo = f"{uuid_str}_FECHADA_{tipo}_"
            sufixo  = ".xlsx"
            if not (base.startswith(prefixo) and base.endswith(sufixo)):
                logger.warning(f"[{tipo}] Ignorando XLSX fora do padrão: {base}")
                continue
            matricula = base[len(prefixo):-len(sufixo)]
            docx_path = os.path.join(diretorio, f"{uuid_str}_FECHADA_{tipo}_{matricula}.docx")
            dxf_path  = os.path.join(diretorio, f"{uuid_str}_FECHADA_{tipo}_{matricula}.dxf")

            ok_xlsx = os.path.exists(excel_path)
            ok_docx = os.path.exists(docx_path)
            ok_dxf  = os.path.exists(dxf_path)
            logger.info(f"[{tipo}/{matricula}] ok: XLSX={ok_xlsx} DOCX={ok_docx} DXF={ok_dxf}")

            if not (ok_xlsx and ok_docx and ok_dxf):
                logger.warning(f"Incompleto para {tipo}/{matricula}. Pulando…")
                continue

            zip_name = f"{uuid_str}_FECHADA_{tipo}_{matricula}.zip"
            zip_path = os.path.join(diretorio, zip_name)

            try:
                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(dxf_path,  os.path.basename(dxf_path))
                    zf.write(docx_path, os.path.basename(docx_path))
                    zf.write(excel_path, os.path.basename(excel_path))
                logger.info(f"✅ ZIP criado: {zip_path}")
            except Exception as e:
                logger.error(f"Erro ao criar ZIP {zip_path}: {e}")

    logger.info("Compactação finalizada")



def main_compactar_arquivos(diretorio_concluido, cidade, uuid_str):
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade, uuid_str)
    logger.info("Compactação finalizada")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Compacta arquivos gerados.")
    parser.add_argument('--diretorio', required=True, help="Diretório dos arquivos concluídos.")
    parser.add_argument('--cidade', required=True, help="Nome formatado da cidade.")
    parser.add_argument('--uuid', required=True, help="UUID da execução atual.")
    args = parser.parse_args()
    logger.info(f"[DEBUG] Conteúdo de {diretorio_concluido}: {os.listdir(diretorio_concluido)}")

    main_compactar_arquivos(args.diretorio, args.cidade, args.uuid)
