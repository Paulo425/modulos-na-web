import argparse
import sys
import codecs
import os
import time

from preparar_arquivos import main_preparo_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

sys.stdout.reconfigure(encoding='utf-8')

# 🔧 Caminho base e pasta pública para arquivos visíveis via web
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
CAMINHO_PUBLICO = os.path.join(BASE_DIR, 'static', 'arquivos')
os.makedirs(CAMINHO_PUBLICO, exist_ok=True)  # Garante criação automática

def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf):
    print("🚀 [main.py] Início da execução principal")

    print("📁 Variáveis de entrada:")
    print(f"   - diretorio_saida: {diretorio_saida}")
    print(f"   - cidade: {cidade}")
    print(f"   - excel: {caminho_excel}")
    print(f"   - dxf: {caminho_dxf}")

    print("\n🔷 Iniciando: Preparo inicial dos arquivos")
    variaveis = main_preparo_arquivos(diretorio_saida, cidade, caminho_excel, caminho_dxf)

    if not variaveis:
        print("❌ [main.py] Erro: O preparo inicial não retornou variáveis.")
        return

    diretorio_final = variaveis["diretorio_final"]
    diretorio_preparado = variaveis["diretorio_preparado"]
    diretorio_concluido = variaveis["diretorio_concluido"]
    arquivo_excel_recebido = variaveis["arquivo_excel_recebido"]
    arquivo_dxf_recebido = variaveis["arquivo_dxf_recebido"]
    caminho_template = variaveis["caminho_template"]

    print("✅ [main.py] Preparo concluído. Variáveis carregadas.")
    print(f"   - diretorio_concluido: {diretorio_concluido}")
    print(f"   - template: {caminho_template}")

    print("\n🔷 Processamento Poligonal Fechada")
    main_poligonal_fechada(
        arquivo_excel_recebido,
        arquivo_dxf_recebido,
        diretorio_preparado,
        diretorio_concluido,
        caminho_template
    )

    print(f"\n📦 [main.py] Chamando compactação no diretório: {diretorio_concluido}")
    main_compactar_arquivos(diretorio_concluido)
    print("✅ [main.py] Compactação finalizada com sucesso!")

    print("\n✅ [main.py] Processo geral concluído com sucesso!")


if __name__ == "__main__":
    print("⚙️ [main.py] Script chamado diretamente via linha de comando")

    parser = argparse.ArgumentParser(description='Executar DECOPA diretamente com parâmetros.')
    parser.add_argument('--diretorio', help='Diretório onde salvar arquivos.')
    parser.add_argument('--cidade', help='Cidade do memorial.')
    parser.add_argument('--excel', help='Caminho do arquivo Excel.')
    parser.add_argument('--dxf', help='Caminho do arquivo DXF.')

    args = parser.parse_args()

    diretorio = args.diretorio
    cidade = args.cidade
    excel = args.excel
    dxf = args.dxf

    # 🔒 Proteção: redireciona diretório inválido do Windows para pasta segura no Render
    if not diretorio or 'C:\\' in diretorio or 'OneDrive' in diretorio:
        diretorio = os.path.join(BASE_DIR, '..', 'CONCLUIDO')

    executar_programa(diretorio, cidade, excel, dxf)
