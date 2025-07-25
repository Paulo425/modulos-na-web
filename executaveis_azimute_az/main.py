import os
import sys
import logging
from datetime import datetime
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
import shutil
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

# ✅ 5. Habilita UTF-8 no console
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass  # Em alguns ambientes, reconfigure não está disponível

def main():
    if len(sys.argv) != 4:
        print("Uso: python main.py <cidade> <caminho_excel> <caminho_dxf>")
        sys.exit(1)

    cidade = sys.argv[1]
    uuid_str = str(uuid.uuid4())[:8]
    cidade_formatada = cidade.replace(" ", "_")  # 🔧 Adicione esta linha
    caminho_excel = sys.argv[2]
    caminho_dxf = sys.argv[3]
    base_dir = os.path.dirname(os.path.abspath(__file__))
    caminho_template = os.path.join(base_dir, "Memorial_modelo_padrao.docx")

    if not os.path.exists(caminho_template):
        print(f"Template '{caminho_template}' não encontrado.")
        sys.exit(1)

    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir, uuid_str)


    main_poligonal_fechada(
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"],
        caminho_template,
        uuid_str
    )


    main_compactar_arquivos(variaveis["diretorio_concluido"], cidade_formatada, uuid_str)

    print("✅ [main.py] Compactação finalizada com sucesso!")

    # 🔁 Copiar ZIPs para static/arquivos e exibir debug
    # ✅ Copiar todos os ZIPs que realmente existem
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
                print(f"📦 ZIP copiado: {arquivo}")
                zips_copiados += 1

        if zips_copiados == 0:
            print("⚠️ Nenhum ZIP encontrado para copiar.")
    except Exception as e:
        print(f"❌ Erro ao copiar ZIPs: {e}")
    
    # 🔎 Verificação final para exibir o nome do ZIP copiado mais recente
    try:
        arquivos_zip = [f for f in os.listdir(pasta_destino) if f.lower().endswith('.zip')]
        if arquivos_zip:
            arquivos_zip.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_destino, x)), reverse=True)
            zip_download = arquivos_zip[0]
            print(f"🔗 ZIP disponível para download: {zip_download}")
    except Exception as e:
        print(f"⚠️ Não foi possível determinar o nome do ZIP: {e}")





if __name__ == "__main__":
    main()
