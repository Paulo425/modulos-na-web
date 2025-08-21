# executaveis/executar_memorial_azimute_jl.py
def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar, dxf_path, excel_path, log_path, sentido_poligonal="horario"):

    import os, math, traceback
    from pathlib import Path
    from datetime import datetime
    import glob,logging

    # Tenta importar o módulo de utilidades JL (suporta os dois nomes que você usa)
    from .memoriais_JL import (
        limpar_dxf_preservando_original,
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,
    )
    #atualizado

    print("[DEBUG] Início da execução do memorial")

    # BASE_DIR absoluto para resolver template corretamente
    BASE_DIR = Path(__file__).resolve().parents[1]
    log.write(f"[JL] sentido_poligonal recebido: {sentido_poligonal}\n")

   
    logger = logging.getLogger(__name__)

    # 0) Deriva UUID e TIPO
    uuid_exec  = os.path.basename(os.path.dirname(os.path.normpath(caminho_salvar)))
    tipo = next((t for t in ["ETE", "REM", "SER", "ACE"] if t in os.path.basename(dxf_path).upper()), None)
    if not tipo:
        logger.error("Tipo não identificado no nome do DXF (esperava conter ETE/REM/SER/ACE).")
        return False

    safe_uuid = sanitize_filename(uuid_exec)[:8]
    safe_tipo = sanitize_filename(tipo)
    safe_mat  = sanitize_filename(matricula)

    # 1) Gerar DXF LIMPO no padrão UUID_TIPO_MATRICULA_LIMPO.dxf (igual DECOPA)
    nome_limpo_dxf    = f"{safe_uuid}_{safe_tipo}_{safe_mat}_LIMPO.dxf"
    caminho_dxf_limpo = os.path.join(caminho_salvar, nome_limpo_dxf)

    dxf_resultado, ponto_az, ponto_inicial = limpar_dxf_e_inserir_ponto_az(dxf_path, caminho_dxf_limpo)
    logger.info(f"[JL] DXF limpo salvo em: {caminho_dxf_limpo}")

    if not ponto_az or not ponto_inicial:
        logger.error("[JL] Não foi possível identificar o ponto Az ou ponto inicial.")
        return False

    # 2) Extrai linhas/arcos do LIMPO e calcula azimute/distância p/ V1 (mesmo miolo do DECOPA)
    doc, linhas, arcos, perimeter_dxf, area_dxf = get_document_info_from_dxf(dxf_resultado)
    if not doc or not linhas:
        logger.error("[JL] Documento DXF inválido ou vazio após limpeza.")
        return False

    msp = doc.modelspace()
    v1 = linhas[0][0]
    distance_az_v1 = calculate_distance(ponto_az, v1)
    azimute_az_v1  = calculate_azimuth(ponto_az, v1)
    logger.info(f"[JL] Azimute: {azimute_az_v1:.2f}°, Distância Az-V1: {distance_az_v1:.2f}m")

    # 3) Gera DXF final + XLSX + (opcional) DOCX no mesmo padrão do DECOPA
    #    Aqui passamos excel_path diretamente (no DECOPA eu procurava FECHADA_*; no JL você já fornece)
    excel_output = create_memorial_descritivo(
        doc=doc,
        msp=msp,
        lines=linhas,
        arcs=arcos,
        proprietario=proprietario,
        matricula=matricula,
        caminho_salvar=caminho_salvar,
        excel_file_path=excel_path,        # <— JL já fornece a planilha apropriada
        ponto_az=ponto_az,
        distance_az_v1=distance_az_v1,
        azimute_az_v1=azimute_az_v1,
        ponto_inicial_real=ponto_inicial,
        tipo=tipo,
        uuid_prefix=safe_uuid,
        sentido_poligonal=sentido_poligonal
    )
    if not excel_output:
        logger.error("[JL] Falha ao gerar memorial descritivo (XLSX/DXF).")
        return False

    # Se você tiver a função de DOCX no JL, mantenha o padrão do nome com UUID_TIPO_MAT
    try:
        output_docx = os.path.join(caminho_salvar, f"{safe_uuid}_{safe_tipo}_{safe_mat}.docx")
        create_memorial_document(
            proprietario=proprietario,
            matricula=matricula,
            descricao=descricao,
            area_terreno="",           # se tiver no JL, passe aqui; senão deixe vazio
            excel_file_path=excel_output,
            template_path=CAMINHO_TEMPLATE_JL,  # defina sua constante no JL
            output_path=output_docx,
            perimeter_dxf=perimeter_dxf,
            area_dxf=area_dxf,
            desc_ponto_Az="",
            Coorde_E_ponto_Az=ponto_az[0],
            Coorde_N_ponto_Az=ponto_az[1],
            azimuth=azimute_az_v1,
            distance=distance_az_v1,
            comarca="",                # preencha se tiver
            RI="",                     # idem
            rua="",                    # idem
            uuid_prefix=safe_uuid
        )
        logger.info(f"[JL] DOCX salvo em: {output_docx}")
    except Exception as e:
        logger.warning(f"[JL] DOCX não gerado ou opcional: {e}")

    # 4) (Opcional) Remover o LIMPO para não confundir compactação
    try:
        if os.path.exists(caminho_dxf_limpo):
            os.remove(caminho_dxf_limpo)
            logger.info(f"[JL] DXF LIMPO removido: {caminho_dxf_limpo}")
    except Exception as e:
        logger.warning(f"[JL] Não foi possível remover DXF LIMPO: {e}")

    return True
