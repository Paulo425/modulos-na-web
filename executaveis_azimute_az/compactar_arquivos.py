import os
import glob
import zipfile

def montar_pacote_zip(diretorio):
    tipos = ["ETE", "REM", "SER", "ACE"]

    for tipo in tipos:
        # Padrões para localizar os arquivos
        padrao_dxf = os.path.join(diretorio, f"{tipo}_Memorial_*.dxf")
        padrao_docx = os.path.join(diretorio, f"{tipo}_Memorial_MAT_*.docx")
        padrao_excel = os.path.join(diretorio, f"{tipo}_Memorial_*.xlsx")

        # Localizar os arquivos
        arquivo_dxf = glob.glob(padrao_dxf)
        arquivo_docx = glob.glob(padrao_docx)
        arquivo_excel = glob.glob(padrao_excel)

        # Conferir arquivos encontrados
        if arquivo_dxf and arquivo_docx and arquivo_excel:
            nome_base = os.path.splitext(os.path.basename(arquivo_docx[0]))[0]
            nome_zip = os.path.join(diretorio, f"{nome_base}.zip")

            print(f"📦 Compactando arquivos para tipo {tipo} em {nome_zip}")

            with zipfile.ZipFile(nome_zip, 'w') as zipf:
                zipf.write(arquivo_dxf[0], os.path.basename(arquivo_dxf[0]))
                zipf.write(arquivo_docx[0], os.path.basename(arquivo_docx[0]))
                zipf.write(arquivo_excel[0], os.path.basename(arquivo_excel[0]))

            print(f"✅ Arquivos do tipo {tipo} compactados com sucesso!")
        else:
            print(f"⚠️ Arquivos incompletos ou não encontrados para o tipo {tipo}")

def main_compactar_arquivos(diretorio_concluido):
    montar_pacote_zip(diretorio_concluido)

