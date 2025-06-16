import os

def buscar_string_em_arquivos(diretorio, termo):
    for pasta_atual, _, arquivos in os.walk(diretorio):
        for nome_arquivo in arquivos:
            if nome_arquivo.endswith(('.html', '.py')):
                caminho_completo = os.path.join(pasta_atual, nome_arquivo)
                with open(caminho_completo, 'r', encoding='utf-8', errors='ignore') as f:
                    for i, linha in enumerate(f, 1):
                        if termo in linha:
                            print(f"{caminho_completo} - linha {i}: {linha.strip()}")

buscar_string_em_arquivos('.', 'gerar_memoriais_azimute_az')
