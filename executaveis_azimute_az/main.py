import os
import sys
import logging
from datetime import datetime
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
import shutil

# ✅ 1. Caminho base
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

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
    cidade_formatada = cidade.replace(" ", "_")  # 🔧 Adicione esta linha
    caminho_excel = sys.argv[2]
    caminho_dxf = sys.argv[3]
    base_dir = os.path.dirname(os.path.abspath(__file__))
    caminho_template = os.path.join(base_dir, "Memorial_modelo_padrao.docx")

    if not os.path.exists(caminho_template):
        print(f"Template '{caminho_template}' não encontrado.")
        sys.exit(1)

    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir)

    main_poligonal_fechada(
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"],
        caminho_template
    )

    main_compactar_arquivos(variaveis["diretorio_concluido"], variaveis["cidade_formatada"])
    print("✅ [main.py] Compactação finalizada com sucesso!")

    # 🔁 Copiar ZIPs para static/arquivos
    try:
        for arquivo in os.listdir(variaveis["diretorio_concluido"]):
            if arquivo.lower().endswith(".zip"):
                origem = os.path.join(variaveis["diretorio_concluido"], arquivo)
                destino = os.path.join(BASE_DIR, "static", "arquivos", arquivo)
                os.makedirs(os.path.dirname(destino), exist_ok=True)
                shutil.copy2(origem, destino)
                print(f"📦 ZIP copiado para pasta pública: {destino}")
    except Exception as e:
        print(f"⚠️ Erro ao copiar ZIP: {e}")


if __name__ == "__main__":
    main()
