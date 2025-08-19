# executaveis/executar_memorial_azimute_jl.py
def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar, dxf_path, excel_path, log_path, sentido_poligonal="horario"):

    import os, math, traceback
    from pathlib import Path
    from datetime import datetime

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

    try:
        Path(caminho_salvar).mkdir(parents=True, exist_ok=True)

        with open(log_path, 'w', encoding='utf-8') as log:
            log.write(f"🟢 LOG iniciado em: {datetime.now()}\n")

            try:
                # 1) DXF limpo dentro do CONCLUIDO
                dxf_limpo_path = os.path.join(caminho_salvar, f"DXF_LIMPO_{matricula}.dxf")
                dxf_limpo_path, ponto_az, ponto_inicial_real = limpar_dxf_preservando_original(
                    dxf_path, dxf_limpo_path, log=log
                )

                # 2) Info do DXF limpo
                doc, lines, arcs, perimeter_dxf, area_dxf, boundary_points = get_document_info_from_dxf(
                    dxf_limpo_path, log=log
                )
                if not doc or not lines:
                    raise ValueError("Erro ao processar o DXF: polilinha fechada não encontrada.")

                if ponto_az is None:
                    ponto_az = (0.0, 0.0)

                v1 = lines[0][0]
                distance = math.hypot(v1[0] - ponto_az[0], v1[1] - ponto_az[1])
                azimuth = math.degrees(math.atan2(v1[0] - ponto_az[0], v1[1] - ponto_az[1]))
                if azimuth < 0:
                    azimuth += 360

                msp = doc.modelspace()

                # 3) Excel + DXF final (ambos em CONCLUIDO)
                excel_output = create_memorial_descritivo(
                    doc=doc,
                    msp=msp,
                    lines=lines,
                    arcs=arcs,
                    proprietario=proprietario,
                    matricula=matricula,
                    caminho_salvar=caminho_salvar,   # <- CONCLUIDO
                    excel_file_path=excel_path,       # planilha de confrontantes enviada
                    ponto_az=ponto_az,
                    distance_az_v1=distance,
                    azimute_az_v1=azimuth,
                    ponto_inicial_real=ponto_inicial_real,
                    boundary_points=boundary_points,
                    log=log,
                    sentido_poligonal=sentido_poligonal,
                )

                # 4) DOCX final em CONCLUIDO, com template ABSOLUTO
                docx_path = os.path.join(caminho_salvar, f"Memorial_MAT_{matricula}.docx")
                template_path = str(BASE_DIR / "templates_doc" / "MODELO_TEMPLATE_DOC_JL_CORRETO.docx")

                create_memorial_document(
                    proprietario,
                    matricula,
                    descricao,
                    excel_file_path=excel_output,
                    template_path=template_path,
                    output_path=docx_path,
                    perimeter_dxf=perimeter_dxf,
                    area_dxf=area_dxf,
                    Coorde_E_ponto_Az=ponto_az[0],
                    Coorde_N_ponto_Az=ponto_az[1],
                    azimuth=azimuth,
                    distance=distance,
                    log=log,
                )

                final_dxf_path = os.path.join(caminho_salvar, f"Memorial_{matricula}.dxf")

                log.write("✅ Processamento finalizado com sucesso.\n")
                return log_path, [excel_output, final_dxf_path, docx_path]

            except Exception as e:
                traceback.print_exc(file=log)
                log.write(f"\n❌ Erro durante execução: {e}\n")
                return log_path, []

    except Exception as e_fora:
        # fallback caso o 'with open' falhe
        with open(log_path, 'a', encoding='utf-8') as log_fallback:
            log_fallback.write(f"\n❌ ERRO GRAVE antes do log principal: {e_fora}\n")
            traceback.print_exc(file=log_fallback)
        return log_path, []
