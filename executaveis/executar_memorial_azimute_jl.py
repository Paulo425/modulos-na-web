def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar,
                         dxf_path, excel_path, log_path, sentido_poligonal="horario"):
    import os, logging
    from pathlib import Path
    from .memoriais_JL import (
        limpar_dxf_basico,
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,   # se não usar DOCX, pode remover
        sanitize_filename,
    )

    # logger + writer com .write(), igual DECOPA
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    class _LogWriter:
        def __init__(self, file_path): self.file_path = file_path
        def write(self, msg):
            try:
                with open(self.file_path, "a", encoding="utf-8") as fh:
                    fh.write(msg if msg.endswith("\n") else msg + "\n")
            except Exception:
                pass
            logger.info(str(msg).strip())
    log = _LogWriter(log_path)
    log.write(f"[JL] sentido_poligonal: {sentido_poligonal}")

    # === Inputs & nomes base (como no DECOPA) ===
    uuid_exec = os.path.basename(os.path.dirname(os.path.normpath(caminho_salvar)))
    safe_uuid = sanitize_filename(uuid_exec)[:8]
    # tipo via nome do DXF (opcional; a create pode deduzir do Excel se quiser)
    tipo = next((t for t in ["ETE", "REM", "SER", "ACE"] if t in os.path.basename(dxf_path).upper()), None)
    safe_tipo = sanitize_filename(tipo) if tipo else "TIPO"
    safe_mat  = sanitize_filename(matricula)

    # === DXF original + DXF limpo (sem Az) ===
    nome_dxf_limpo    = f"{safe_uuid}_{safe_tipo}_{safe_mat}_LIMPO.dxf"
    caminho_dxf_limpo = os.path.join(caminho_salvar, nome_dxf_limpo)
    ret_clean = limpar_dxf_basico(dxf_path, caminho_dxf_limpo, log=log)
    if not (isinstance(ret_clean, tuple) and len(ret_clean) == 3):
        log.write(f"[ERRO] limpar_dxf_basico retornou {type(ret_clean)} -> {ret_clean!r}")
        return False
    dxf_resultado, ponto_inicial_real, _resumo = ret_clean

    # === Parse do DXF limpo, igual DECOPA ===
    ret_info = get_document_info_from_dxf(dxf_resultado)
    if not (isinstance(ret_info, tuple) and len(ret_info) == 5):
        log.write(f"[ERRO] get_document_info_from_dxf retornou {type(ret_info)} -> {ret_info!r}")
        return False
    doc, linhas, arcos, perimeter_dxf, area_dxf = ret_info
    if not doc or not linhas:
        log.write("[ERRO] DXF inválido ou sem linhas após limpeza.")
        return False
    msp = doc.modelspace()

    # === Chamada idêntica ao DECOPA (sem Az) ===
    excel_output = create_memorial_descritivo(
        doc=doc,
        msp=msp,
        lines=linhas,
        arcs=arcos,
        proprietario=proprietario,
        matricula=matricula,
        caminho_salvar=caminho_salvar,
        excel_file_path=excel_path,        # DECOPA também passa a planilha
        ponto_inicial_real=ponto_inicial_real,
        log=log,
        sentido_poligonal=sentido_poligonal
        # OBS: não passamos 'tipo' nem 'uuid_prefix' aqui; sua create pode deduzir.
    )
    if not excel_output:
        log.write("[ERRO] Falha ao gerar memorial descritivo (XLSX/DXF).")
        return False

    # === (Opcional) DOCX no padrão DECOPA ===
    try:
        output_docx = os.path.join(caminho_salvar, f"{safe_uuid}_{safe_tipo}_{safe_mat}.docx")
        create_memorial_document(
            proprietario=proprietario,
            matricula=matricula,
            descricao=descricao,
            area_terreno="",                 # preencha se tiver
            excel_file_path=excel_output,
            template_path=CAMINHO_TEMPLATE_JL,  # defina sua constante
            output_path=output_docx,
            perimeter_dxf=perimeter_dxf,
            area_dxf=area_dxf,
            # Nada de Az aqui
        )
        log.write(f"[JL] DOCX salvo em: {output_docx}")
    except Exception as e:
        log.write(f"[JL] DOCX opcional não gerado: {e}")

    return True
