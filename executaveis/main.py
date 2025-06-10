import argparse
import sys
import codecs
from preparar_arquivos import main_preparo_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos
import time
import os

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def executar_programa(diretorio_saida, cidade, caminho_excel, caminho_dxf):
    print("\n🔷 Iniciando: Preparo inicial dos arquivos")

    variaveis = main_preparo_arquivos(
        diretorio_saida, cidade, caminho_excel, caminho_dxf)

    if not variaveis:
        print("❌ Erro: O preparo inicial não retornou variáveis.")
        return

    diretorio_final = variaveis["diretorio_final"]
    diretorio_preparado = variaveis["diretorio_preparado"]
    diretorio_concluido = variaveis["diretorio_concluido"]
    arquivo_excel_recebido = variaveis["arquivo_excel_recebido"]
    arquivo_dxf_recebido = variaveis["arquivo_dxf_recebido"]
    caminho_template = variaveis["caminho_template"]

    print("\n🔷 Processamento Poligonal Fechada")
    main_poligonal_fechada(arquivo_excel_recebido, arquivo_dxf_recebido, diretorio_preparado, diretorio_concluido, caminho_template)

    print("\n🔷 Compactação final dos arquivos")
    main_compactar_arquivos(diretorio_concluido)

    print("\n✅ Processo concluído com sucesso!")


if __name__ == "__main__":
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

    # 🔒 Proteção: redireciona diretório inválido do Windows para uma pasta segura no Render
    if not diretorio or 'C:\\' in diretorio or 'OneDrive' in diretorio:
        diretorio = os.path.join(BASE_DIR, '..', 'CONCLUIDO')

    executar_programa(diretorio, cidade, excel, dxf)
