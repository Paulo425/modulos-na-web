

def calculate_azimuth(p1, p2):
    delta_x = p2[0] - p1[0]
    delta_y = p2[1] - p1[1]
    azimuth_rad = math.atan2(delta_x, delta_y)
    azimuth_deg = math.degrees(azimuth_rad)
    if azimuth_deg < 0:
        azimuth_deg += 360
    return azimuth_deg


def calculate_distance(point1, point2):
    dx = point2[0] - point1[0]
    dy = point2[1] - point1[1]
    return math.sqrt(dx**2 + dy**2)

def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar,
                         dxf_path, excel_path, log_path, sentido_poligonal="horario"):
    import os, logging, glob, math
    from pathlib import Path
    from .memoriais_JL import (
        limpar_dxf_e_inserir_ponto_az,
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,   # se não usar DOCX, pode remover
        sanitize_filename,
    )
    diretorio_concluido=caminho_salvar
    uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(diretorio_concluido)))
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

    nome_dxf = os.path.basename(dxf_path).upper()
    tipo = next((t for t in ["ETE", "REM", "SER", "ACE"] if t in nome_dxf), None)

    if not tipo:
        msg = "❌ Tipo do projeto não identificado no nome do DXF."
        print(msg)
        logger.warning(msg)
        return

    print(f"📁 Tipo identificado: {tipo}")
    logger.info(f"Tipo identificado: {tipo}")

    # Novo padrão correto
    padrao = os.path.join(diretorio_concluido, f"FECHADA_*_{tipo}.xlsx")
    lista_encontrada = glob.glob(padrao)

    if not lista_encontrada:
        print(f"❌ Arquivo de confrontantes esperado não encontrado para tipo {tipo}")
        logger.warning(f"Nenhuma planilha FECHADA_*_{tipo}.xlsx encontrada.")
        return

    excel_confrontantes = lista_encontrada[0]
    print(f"✅ Confrontante carregado: {excel_confrontantes}")
    logger.info(f"Planilha de confrontantes usada: {excel_confrontantes}")    

    # === Inputs & nomes base (como no DECOPA) ===
    uuid_exec = os.path.basename(os.path.dirname(diretorio_concluido))
    safe_uuid = sanitize_filename(uuid_exec)[:8]  # use a mesma variável que você loga em "[DEBUG] UUID recebido"
    safe_tipo = sanitize_filename(tipo)           # "ETE", "REM", etc.
    safe_mat  = sanitize_filename(matricula)

    nome_limpo_dxf   = f"{safe_uuid}_{safe_tipo}_{safe_mat}_LIMPO.dxf"
    caminho_dxf_limpo = os.path.join(diretorio_concluido, nome_limpo_dxf)

    dxf_resultado, ponto_az, ponto_inicial = limpar_dxf_e_inserir_ponto_az(dxf_path, caminho_dxf_limpo)
    logger.info(f"DXF limpo salvo em: {caminho_dxf_limpo}")

    # (opcional) remover arquivo legado sem UUID
    legado = os.path.join(diretorio_concluido, f"DXF_LIMPO_{safe_mat}.dxf")
    if os.path.exists(legado):
        try:
            os.remove(legado)
            logger.info(f"DXF limpo legado removido: {legado}")
        except Exception as e:
            logger.warning(f"Não foi possível remover legado {legado}: {e}")
    # <<< PATCH
 
    if not ponto_az or not ponto_inicial:
        msg = "❌ Não foi possível identificar o ponto Az ou inicial."
        print(msg)
        logger.error(msg)
        return

    doc, linhas, arcos, perimeter_dxf, area_dxf = get_document_info_from_dxf(dxf_resultado)
    if not doc or not linhas:
        msg = "❌ Documento DXF inválido ou vazio."
        print(msg)
        logger.error(msg)
        return

    msp = doc.modelspace()
    v1 = linhas[0][0]
    distance_az_v1 = calculate_distance(ponto_az, v1)
    azimute_az_v1 = calculate_azimuth(ponto_az, v1)
    azimuth = calculate_azimuth(ponto_az, v1)
    distance = math.hypot(v1[0] - ponto_az[0], v1[1] - ponto_az[1])

    # === Chamada idêntica ao DECOPA (sem Az) ===
    excel_output = create_memorial_descritivo(
        doc=doc,
        msp=msp,
        lines=linhas,
        proprietario=proprietario,
        matricula=matricula,
        caminho_salvar=caminho_salvar,
        arcs=arcos, 
        excel_file_path=excel_path,
        ponto_az=None,
        distance_az_v1=distance_az_v1,
        azimute_az_v1=azimute_az_v1,        # DECOPA também passa a planilha
        ponto_inicial_real=ponto_inicial,
        tipo=tipo,
        uuid_prefix=uuid_prefix,
        diretorio_concluido=diretorio_concluido,
        sentido_poligonal=sentido_poligonal
        # OBS: não passamos 'tipo' nem 'uuid_prefix' aqui; sua create pode deduzir.
    )

    try:
        if os.path.exists(caminho_dxf_limpo):
            os.remove(caminho_dxf_limpo)
            logger.info(f"DXF LIMPO removido após gerar DXF final: {caminho_dxf_limpo}")
    except Exception as e:
        logger.warning(f"Não foi possível remover DXF LIMPO: {e}")

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
