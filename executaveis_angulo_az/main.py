import os
import sys
import logging
from datetime import datetime
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
import shutil
import uuid
import glob  # ← FALTAVA


# ✅ 1. Caminho base
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


# ✅ 2. Pastas públicas
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)

# ✅ 3. Pasta de logs
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(LOG_DIR, f"log_ANGULOAZ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


# ✅ 4. Configura logger
# Configuração do logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler(log_path, encoding='utf-8')
console_handler = logging.StreamHandler(sys.stdout)

formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)


# ✅ 5. Habilita UTF-8 no console
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass  # Em alguns ambientes, reconfigure não está disponível

def main():
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print("Uso: python main.py <cidade> <caminho_excel> <caminho_dxf> [sentido_poligonal]")
        sys.exit(1)

    cidade = sys.argv[1]
    uuid_str = str(uuid.uuid4())[:8]
    cidade_formatada = cidade.replace(" ", "_")
    caminho_excel = sys.argv[2]
    caminho_dxf = sys.argv[3]
    sentido_poligonal = sys.argv[4] if len(sys.argv) == 5 else 'horario'
    logger.info(f"Sentido poligonal recebido no main.py: {sentido_poligonal}")
    caminho_template = os.path.join(BASE_DIR, "templates_doc", "Memorial_modelo_padrao.docx")



    if not os.path.exists(caminho_template):
        logger.error(f"Template não encontrado em '{caminho_template}'.")
        sys.exit(1)

    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, BASE_DIR, uuid_str)

    if not variaveis:
        logger.error("Erro ao preparar arquivos. Encerrando execução.")
        sys.exit(1)

    logger.info("✅ Preparação dos arquivos concluída.")



    main_poligonal_fechada(
        uuid_str,
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"],
        caminho_template,
        sentido_poligonal
  
    )

    logger.info("✅ Processamento da poligonal fechada concluído.")
    # Checkpoint: antes de compactar, confirme se existem artefatos para pelo menos um TIPO
    tipos = ["ETE", "REM", "SER", "ACE"]
    tem_algum = False
    for _tipo in tipos:
        pad_xlsx = os.path.join(variaveis["diretorio_concluido"], f"{uuid_str}_FECHADA_{_tipo}_*.xlsx")
        if glob.glob(pad_xlsx):
            tem_algum = True
            break

    if not tem_algum:
        logger.error("❌ Nenhum XLSX FECHADA gerado em %s. Compactação não terá o que zipar.",
                    variaveis["diretorio_concluido"])

    main_compactar_arquivos(variaveis["diretorio_concluido"], cidade_formatada, uuid_str)

    logger.info("✅ Compactação concluída com sucesso.")


    # 🔁 Copiar ZIPs para static/arquivos e exibir debug
    # ✅ Copiar todos os ZIPs que realmente existem
    # 🔁 Copiar ZIPs para static/arquivos
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
                logger.info(f"📦 ZIP copiado: {arquivo}")
                zips_copiados += 1

        if zips_copiados == 0:
            logger.warning("⚠️ Nenhum ZIP encontrado para copiar.")
    except Exception as e:
        logger.error(f"❌ Erro ao copiar ZIPs: {e}")

    # 🔎 Verificação final - ZIP mais recente
    try:
        arquivos_zip = [
            f for f in os.listdir(pasta_destino)
            if f.lower().endswith('.zip') and uuid_str in f
        ]
        if arquivos_zip:
            arquivos_zip.sort(
                key=lambda x: os.path.getmtime(os.path.join(pasta_destino, x)),
                reverse=True
            )
            zip_download = arquivos_zip[0]
            logger.info(f"🔗 ZIP disponível para download: {zip_download}")
        else:
            logger.warning("⚠️ Nenhum ZIP disponível para download.")
    except Exception as e:
        logger.error(f"⚠️ Não foi possível determinar o nome do ZIP: {e}")





if __name__ == "__main__":
    main()
#ATUALIZADO