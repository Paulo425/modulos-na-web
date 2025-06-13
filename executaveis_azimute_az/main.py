import os
import sys
import time
from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

def main():
    if len(sys.argv) != 4:
        print("Uso: python main.py <cidade> <caminho_excel> <caminho_dxf>")
        sys.exit(1)

    cidade = sys.argv[1]
    caminho_excel = sys.argv[2]
    caminho_dxf = sys.argv[3]
    base_dir = os.getcwd()
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

    main_compactar_arquivos(variaveis["diretorio_concluido"])

if __name__ == "__main__":
    main()
