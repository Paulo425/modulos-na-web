import os
import glob
import zipfile
import re
import logging
from datetime import datetime
import shutil
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

def montar_pacote_zip(diretorio, cidade):
    print("\n📦 [compactar] Iniciando montagem dos pacotes ZIP")
    logger.info("Iniciando montagem dos pacotes ZIP")

    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        print(f"🔍 Buscando arquivos do tipo: {tipo}")
        logger.info(f"Buscando arquivos do tipo: {tipo}")

        arquivos_dxf = glob.glob(os.path.join(diretorio, f"{tipo}_*.dxf"))
        arquivos_docx = glob.glob(os.path.join(diretorio, f"{tipo}_*.docx"))
        arquivos_excel = glob.glob(os.path.join(diretorio, f"{tipo}_*.xlsx"))

        print(f"   - DXF encontrados: {len(arquivos_dxf)}")
        print(f"   - DOCX encontrados: {len(arquivos_docx)}")
        print(f"   - XLSX encontrados: {len(arquivos_excel)}")

        logger.info(f"DXF={len(arquivos_dxf)} | DOCX={len(arquivos_docx)} | XLSX={len(arquivos_excel)}")

        # Coletar todas as matrículas
        matriculas = set()
        for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
            nome_arquivo = os.path.basename(arq)
            match = re.search(rf"{tipo}[_ ]?(\d+[.,]?\d*)", nome_arquivo)
            if match:
                matricula = match.group(1).replace(",", ".").replace(" ", "")
                if not "." in matricula and len(matricula) > 3:
                    matricula = f"{matricula[:-3]}.{matricula[-3:]}"
                matriculas.add(matricula)

        for matricula in matriculas:
            print(f"\n🔢 Processando matrícula: {matricula}")
            logger.info(f"Processando matrícula: {matricula}")

            arq_dxf = [a for a in arquivos_dxf if matricula in a]
            arq_docx = [a for a in arquivos_docx if matricula in a]
            arq_excel = [a for a in arquivos_excel if matricula in a]

            if arq_dxf and arq_docx and arq_excel:
                cidade_sanitizada = cidade.replace(" ", "_")
                uuid_prefix = os.path.basename(diretorio)
                nome_zip = os.path.join(diretorio, f"{uuid_prefix}_{cidade_sanitizada}_{tipo}_{matricula}.zip")

                try:
                    uuid_prefix = os.path.basename(diretorio)  # EX: '8f92ac18'

                    nome_zip = os.path.join(diretorio, f"{uuid_prefix}_{cidade_sanitizada}_{tipo}_{matricula}.zip")
                    temp_dir = os.path.join(diretorio, "TEMP_ZIP")
                    os.makedirs(temp_dir, exist_ok=True)

                    caminho_docx = os.path.join(temp_dir, f"{uuid_prefix}_{tipo}_{matricula}.docx")
                    caminho_dxf  = os.path.join(temp_dir, f"{uuid_prefix}_{tipo}_{matricula}.dxf")
                    caminho_xlsx = os.path.join(temp_dir, f"{uuid_prefix}_{tipo}_{matricula}.xlsx")

                    shutil.copy2(arq_docx[0], caminho_docx)
                    shutil.copy2(arq_dxf[0], caminho_dxf)
                    shutil.copy2(arq_excel[0], caminho_xlsx)

                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(caminho_docx, os.path.basename(caminho_docx))
                        zipf.write(caminho_dxf,  os.path.basename(caminho_dxf))
                        zipf.write(caminho_xlsx, os.path.basename(caminho_xlsx))

                    shutil.copy2(nome_zip, caminho_debug_zip)

                    print(f"✅ ZIP criado com sucesso: {nome_zip}")
                    logger.info(f"ZIP criado: {nome_zip} e copiado para: {caminho_debug_zip}")

                    # 🔁 Limpar pasta TEMP_ZIP após uso
                    try:
                        shutil.rmtree(temp_dir)
                        print(f"🧹 Pasta temporária removida: {temp_dir}")
                        logger.info(f"Pasta temporária removida: {temp_dir}")
                    except Exception as e:
                        print(f"⚠️ Falha ao remover pasta temporária: {e}")
                        logger.warning(f"Falha ao remover pasta temporária {temp_dir}: {e}")

                except Exception as e:
                    print(f"❌ Erro ao criar ZIP: {e}")
                    logger.error(f"Erro ao criar ZIP {nome_zip}: {e}")




def main_compactar_arquivos(diretorio_concluido, cidade_formatada):
    print(f"\n📦 Iniciando compactação no diretório: {diretorio_concluido}")
    logger.info(f"Iniciando compactação no diretório: {diretorio_concluido}")
    montar_pacote_zip(diretorio_concluido, cidade_formatada)
    print("✅ Compactação finalizada")
    logger.info("Compactação finalizada")

if __name__ == "__main__":
    import argparse

    TMP_CONCLUIDO = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')
    parser = argparse.ArgumentParser(description="Compacta arquivos gerados em ZIP.")
    parser.add_argument('--diretorio', default=TMP_CONCLUIDO, help="Diretório com os arquivos a compactar.")
    args = parser.parse_args()

    main_compactar_arquivos(args.diretorio)
