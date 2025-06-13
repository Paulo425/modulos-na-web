# Versão ajustada para ambiente Render (sem pythoncom)
# ... restante do código mantido ...

# [FUNÇÕES AQUI]

    # Salvar o arquivo DOCX com as alterações
    try:
        docx_output_path = os.path.normpath(os.path.join(caminho_salvar, f"{tipo}_Memorial_MAT_{matricula}.docx"))
        doc_word.save(docx_output_path)
        print(f"Memorial descritivo salvo em: {docx_output_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo DOCX: {e}")
