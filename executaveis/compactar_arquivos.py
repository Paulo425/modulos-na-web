import os
import glob
import zipfile
import re

def montar_pacote_zip(diretorio):
    print("🟨 Conteúdo da pasta antes da compactação:", os.listdir(diretorio))

    tipos = ["ETE", "REM", "SER", "ACE"]

    # Expressão regular para extrair matrícula
    import re
    matricula_regex = re.compile(r"_([0-9]+\.[0-9]+)")

    for tipo in tipos:
        arquivos_dxf = glob.glob(os.path.join(diretorio, f"{tipo}_Memorial_*.dxf"))
        arquivos_docx = glob.glob(os.path.join(diretorio, f"{tipo}_Memorial_*.docx"))
        arquivos_excel = glob.glob(os.path.join(diretorio, f"{tipo}_Memorial_*.xlsx"))

        # Extrair todas as matrículas
        matriculas = set()
        for arq in arquivos_docx + arquivos_dxf + arquivos_excel:
            match = matricula_regex.search(arq)
            if match:
                matriculas.add(match.group(1))

        for matricula in matriculas:
            arq_dxf = [a for a in arquivos_dxf if matricula in a]
            arq_docx = [a for a in arquivos_docx if matricula in a]
            arq_excel = [a for a in arquivos_excel if matricula in a]

            if arq_dxf and arq_docx and arq_excel:
                nome_zip = os.path.join(diretorio, f"{tipo}_Memorial_MAT_{matricula}.zip")
                print(f"📦 Compactando arquivos para {tipo}, matrícula {matricula}:")
                print(f"📂 Zip será salvo como: {nome_zip}")

                with zipfile.ZipFile(nome_zip, 'w') as zipf:
                    zipf.write(arq_dxf[0], os.path.basename(arq_dxf[0]))
                    zipf.write(arq_docx[0], os.path.basename(arq_docx[0]))
                    zipf.write(arq_excel[0], os.path.basename(arq_excel[0]))

                print(f"✅ Compactado com sucesso: {nome_zip}")
                print(f"🔍 Nome do ZIP final criado: {os.path.basename(nome_zip)}")

            else:
                print(f"⚠️ Arquivos incompletos para {tipo}, matrícula {matricula}")
                
def main_compactar_arquivos(diretorio_concluido):
    montar_pacote_zip(diretorio_concluido)

# 🔽 Suporte para execução via linha de comando (opcional)
if __name__ == "__main__":
    import argparse

    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    TMP_CONCLUIDO = os.path.join(BASE_DIR, 'tmp', 'CONCLUIDO')

    parser = argparse.ArgumentParser(description="Compacta arquivos gerados em ZIP.")
    parser.add_argument('--diretorio', default=TMP_CONCLUIDO, help="Diretório com os arquivos a compactar.")

    args = parser.parse_args()
    main_compactar_arquivos(args.diretorio)
