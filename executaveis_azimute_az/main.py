import os
import sys
import time
from preparar_arquivos import main_preparo_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

if hasattr(sys, '_MEIPASS'):
    diretorio_atual = sys._MEIPASS
else:
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

caminho_template = os.path.join(diretorio_atual, "Memorial_modelo_padrao.docx")

# Verificação essencial antes de continuar o script:
if not os.path.exists(caminho_template):
    print(f"ERRO: O arquivo '{{caminho_template}}' não foi encontrado!")
    input("Pressione ENTER para sair...")
    exit()

def main():
    print("\n🔷 Iniciando: Preparo inicial dos arquivos")
    variaveis = main_preparo_arquivos()
    if not variaveis:
        print("❌ Erro: O preparo inicial não retornou variáveis.")
        return

    time.sleep(2)
    diretorio_final = variaveis["diretorio_final"]
    diretorio_recebido_carlos = variaveis["diretorio_recebido_carlos"]
    diretorio_preparado = variaveis["diretorio_preparado"]
    diretorio_concluido = variaveis["diretorio_concluido"]
    arquivo_excel_recebido = variaveis["arquivo_excel_recebido"]
    arquivo_dxf_recebido = variaveis["arquivo_dxf_recebido"]

    print("\n🔷 Iniciando: Processamento Poligonal Fechada")
    main_poligonal_fechada(arquivo_excel_recebido, arquivo_dxf_recebido, diretorio_preparado, diretorio_concluido, caminho_template)
    time.sleep(2)

    print("\n🔷 Iniciando: Compactação final dos arquivos")
    main_compactar_arquivos(diretorio_concluido)

    print("\n✅ Processo completo concluído com sucesso!")

if __name__ == "__main__":
    main()
