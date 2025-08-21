# executaveis/executar_memorial_azimute_jl.py
def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar,
                         dxf_path, excel_path, log_path, sentido_poligonal="horario"):

    import os, math, traceback, logging
    from pathlib import Path
    from .memoriais_JL import (
        limpar_dxf_basico,
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,
        sanitize_filename,
    )

    print("[DEBUG] Início da execução do memorial")

    BASE_DIR = Path(__file__).resolve().parents[1]

    # ✅ Logger com writer .write()
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    class _LogWriter:
        def __init__(self, file_path):
            self.file_path = file_path
        def write(self, msg):
            try:
                with open(self.file_path, "a", encoding="utf-8") as fh:
                    fh.write(msg if msg.endswith("\n") else msg + "\n")
            except Exception:
                pass
            logger.info(msg.strip())

    log = _LogWriter(log_path)
    log.write(f"[JL] sentido_poligonal recebido: {sentido_poligonal}")

    # 0) Deriva UUID e TIPO
    uuid_exec  = os.path.basename(os.path.dirname(os.path.normpath(caminho_salvar)))
    tipo = next((t for t in ["ETE", "REM", "SER", "ACE"] if t in os.path.basename(dxf_path).upper()), None)
    if not tipo:
        logger.error("Tipo não identificado no nome do DXF (esperava conter ETE/REM/SER/ACE).")
        return False

    safe_uuid = sanitize_filename(uuid_exec)[:8]
    safe_tipo = sanitize_filename(tipo)
    safe_mat  = sanitize_filename(matricula)

    # 1) Limpeza básica de DXF (sem Az)
    nome_limpo_dxf    = f"{safe_uuid}_{safe_tipo}_{safe_mat}_LIMPO.dxf"
    caminho_dxf_limpo = os.path.join(caminho_salvar, nome_limpo_dxf)

    ret_clean = limpar_dxf_basico(dxf_path, caminho_dxf_limpo, log=log)
    if not isinstance(ret_clean, tuple) or len(ret_clean) != 3:
        logger.error(f"[JL] limpar_dxf_basico retornou {type(ret_clean)} -> {ret_clean!r}")
        return False
    dxf_resultado, ponto_inicial_real, resumo_limpeza = ret_clean
    logger.info(f"[JL] DXF limpo em: {dxf_resultado} | resumo={resumo_limpeza}")

    # 2) Extrai linhas/arcos do DXF limpo e perímetro/área
    ret_info = get_document_info_from_dxf(dxf_resultado)
    if not isinstance(ret_info, tuple) or len(ret_info) != 5:
        logger.error(f"[JL] get_document_info_from_dxf retornou {type(ret_info)} -> {ret_info!r}")
        return False
    
    
    
    # 1) Limpeza básica de DXF (sem Az)
    ret_clean = limpar_dxf_basico(dxf_path, caminho_dxf_limpo, log=log)

    # 🔒 Guard: evita 'cannot unpack non-iterable bool object'
    if not isinstance(ret_clean, tuple) or len(ret_clean) != 3:
        logger.error(f"[JL] limpar_dxf_basico retornou {type(ret_clean)} -> {ret_clean!r}")
        return False

    dxf_resultado, ponto_inicial_real, resumo_limpeza = ret_clean
    logger.info(f"[JL] DXF limpo em: {dxf_resultado} | resumo={resumo_limpeza}")

    # 2) Extrai linhas/arcos do DXF limpo e perímetro/área
    ret_info = get_document_info_from_dxf(dxf_resultado)

    # 🔒 Guard: evita 'cannot unpack non-iterable bool object'
    if not isinstance(ret_info, tuple) or len(ret_info) != 5:
        logger.error(f"[JL] get_document_info_from_dxf retornou {type(ret_info)} -> {ret_info!r}")
        return False

    doc, linhas, arcos, perimeter_dxf, area_dxf = ret_info


    if not doc or not linhas:
        logger.error("[JL] Documento DXF inválido ou sem linhas após limpeza.")
        return False

    msp = doc.modelspace()

    # 3) Gera DXF final + XLSX
    excel_output = create_memorial_descritivo(
        doc=doc,
        msp=msp,
        lines=linhas,
        arcs=arcos,
        proprietario=proprietario,
        matricula=matricula,
        caminho_salvar=caminho_salvar,
        excel_file_path=excel_path,           # JL já fornece a planilha apropriada
        ponto_inicial_real=ponto_inicial_real,  # ✅ opcional, ajuda a fixar V1
        tipo=tipo,
        uuid_prefix=safe_uuid,
        sentido_poligonal=sentido_poligonal,
        log=log
    )
    if not excel_output:
        logger.error("[JL] Falha ao gerar memorial descritivo (XLSX/DXF).")
        return False

    # 4) DOCX (sem campos de Az; assinatura já deve aceitar Az como opcional)
    try:
        output_docx = os.path.join(caminho_salvar, f"{safe_uuid}_{safe_tipo}_{safe_mat}.docx")
        create_memorial_document(
            proprietario=proprietario,
            matricula=matricula,
            descricao=descricao,
            area_terreno="",                 # se houver, você preenche
            excel_file_path=excel_output,
            template_path=CAMINHO_TEMPLATE_JL,  # sua constante
            output_path=output_docx,
            perimeter_dxf=perimeter_dxf,
            area_dxf=area_dxf,
            # todos os campos de Az removidos/None por padrão
        )
        logger.info(f"[JL] DOCX salvo em: {output_docx}")
    except Exception as e:
        logger.warning(f"[JL] DOCX não gerado ou opcional: {e}")

    # 5) (Opcional) Remover *_LIMPO.dxf se quiser evitar confusão
    try:
        if os.path.exists(caminho_dxf_limpo):
            os.remove(caminho_dxf_limpo)
            logger.info(f"[JL] DXF LIMPO removido: {caminho_dxf_limpo}")
    except Exception as e:
        logger.warning(f"[JL] Não foi possível remover DXF LIMPO: {e}")

    return True

