import os
import traceback
import math
from executaveis.memoriais_JL import (
    limpar_dxf_e_inserir_ponto_az,
    get_document_info_from_dxf,
    create_memorial_descritivo,
    create_memorial_document
)

def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar, dxf_path, excel_path):
    log_path = os.path.join(caminho_salvar, 'execucao.log')
    with open(log_path, 'w', encoding='utf-8') as log:
        try:
            # 🔹 Gerar DXF limpo e obter ponto Az
            dxf_limpo_path = os.path.join(caminho_salvar, f"DXF_LIMPO_{matricula}.dxf")
            dxf_limpo_path, ponto_az = limpar_dxf_e_inserir_ponto_az(dxf_path, dxf_limpo_path)

            # 🔹 Extrair informações do DXF
            doc, lines, arcs, perimeter_dxf, area_dxf = get_document_info_from_dxf(dxf_limpo_path, log=log)
            if not doc or not ponto_az:
                raise ValueError("Erro ao processar o DXF ou ponto Az não encontrado.")

            # 🔹 Calcular azimute e distância do ponto Az ao V1
            v1 = lines[0][0]
            distance = math.hypot(v1[0] - ponto_az[0], v1[1] - ponto_az[1])
            azimuth = math.degrees(math.atan2(v1[0] - ponto_az[0], v1[1] - ponto_az[1]))
            if azimuth < 0:
                azimuth += 360

            msp = doc.modelspace()

            # 🔹 Criar Excel
            excel_output = create_memorial_descritivo(
                doc=doc, msp=msp, lines=lines, arcs=arcs,
                proprietario=proprietario, matricula=matricula,
                caminho_salvar=caminho_salvar, excel_file_path=excel_path,
                ponto_az=ponto_az, distance_az_v1=distance,
                azimute_az_v1=azimuth, log=log
            )

            # 🔹 Criar DOCX
            docx_path = os.path.join(caminho_salvar, f"Memorial_MAT_{matricula}.docx")
            template_path = os.path.join("templates_doc", "MODELO_TEMPLATE_DOC_JL_CORRETO.docx")

            create_memorial_document(
                proprietario, matricula, descricao,
                excel_file_path=excel_output,
                template_path=template_path,
                output_path=docx_path,
                perimeter_dxf=perimeter_dxf,
                area_dxf=area_dxf,
                Coorde_E_ponto_Az=ponto_az[0],
                Coorde_N_ponto_Az=ponto_az[1],
                azimuth=azimuth,
                distance=distance,
                log=log
            )

            # 🔹 Copia final do DXF com anotações
            final_dxf_path = os.path.join(caminho_salvar, f"Memorial_{matricula}.dxf")

            # ✅ Retorna o log e a lista dos arquivos de saída
            return log_path, [excel_output, final_dxf_path, docx_path]

        except Exception as e:
            traceback.print_exc(file=log)
            log.write(f"\nErro: {e}\n")
            return log_path, []
