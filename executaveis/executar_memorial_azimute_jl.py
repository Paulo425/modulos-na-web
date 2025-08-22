

# --- IMPORTS DE MÓDULO (fora da função; necessários para helpers e cálculos) ---
import os
import math
import traceback
import re
from pathlib import Path




# fallback simples caso sanitize_filename não exista no módulo
def _sanitize_filename(s):
    return re.sub(r'[^0-9A-Za-z._-]+', '_', str(s)).strip('_')

try:
    # se existir no seu módulo, usa o oficial
    from .memoriais_JL import sanitize_filename as _sanitize_filename
except Exception:
    pass


def calculate_azimuth(p1, p2):
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    azimuth_rad = math.atan2(dx, dy)
    azimuth_deg = math.degrees(azimuth_rad)
    if azimuth_deg < 0:
        azimuth_deg += 360
    return azimuth_deg


def calculate_distance(p1, p2):
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    return math.sqrt(dx*dx + dy*dy)


def executar_memorial_jl(proprietario, matricula, descricao, caminho_salvar,
                         dxf_path, excel_path, log_path, sentido_poligonal="horario"):
    import logging
    # Importa utilitários do seu módulo JL
    from .memoriais_JL import (
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,
        limpar_dxf_e_inserir_ponto_az,   # ← exatamente como no DECOPA
    )


    # ===== logger + writer simples (grava e também manda pro logger do módulo) =====
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    class _LogWriter:
        def __init__(self, file_path):
            self.file_path = file_path
            # abre o log zerando no início p/ facilitar o diagnóstico da execução corrente
            try:
                with open(self.file_path, "w", encoding="utf-8") as fh:
                    fh.write(f"🟢 LOG iniciado\n")
            except Exception:
                pass

        def write(self, msg):
            try:
                with open(self.file_path, "a", encoding="utf-8") as fh:
                    fh.write(msg if str(msg).endswith("\n") else str(msg) + "\n")
            except Exception:
                pass
            try:
                logger.info(str(msg).strip())
            except Exception:
                pass

    log = _LogWriter(log_path)
    log.write(f"[JL] Iniciando executar_memorial_jl | sentido_poligonal={sentido_poligonal}")

    try:
        # === garante que a pasta CONCLUIDO exista ===
        Path(caminho_salvar).mkdir(parents=True, exist_ok=True)

        # === UUID prefix (basename do pai da pasta CONCLUIDO) ===
        uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(caminho_salvar))) or "JL"
        safe_mat = _sanitize_filename(matricula)

        # === cria caminho do DXF limpo ===
        nome_limpo_dxf = f"{uuid_prefix}_JL_{safe_mat}_LIMPO.dxf"
        caminho_dxf_limpo = os.path.join(caminho_salvar, nome_limpo_dxf)

        # === limpeza do DXF (aceita as duas variantes) ===
        # === limpeza do DXF — padrão DECOPA ===
        log.write("[JL] Limpando DXF com: limpar_dxf_e_inserir_ponto_az")
        res = limpar_dxf_e_inserir_ponto_az(dxf_path, caminho_dxf_limpo)

        # Blindagem do retorno (evita 'cannot unpack non-iterable bool object')
        if not isinstance(res, tuple) or len(res) != 3:
            log.write(f"[ERRO] limpar_dxf_e_inserir_ponto_az retornou {type(res).__name__}: {res!r}")
            return False

        dxf_resultado, ponto_az, ponto_inicial = res
        # manter compatibilidade com o restante do código
        ponto_inicial_real = ponto_inicial


        # Se não houver ponto_az, define neutro (queremos manter assinatura da create)
        if not ponto_az:
            ponto_az = (0.0, 0.0)

        # === leitura das entidades do DXF ===
        try:
            res_info = get_document_info_from_dxf(dxf_resultado)
        except TypeError:
            # Algumas versões aceitam log=..., tenta novamente
            res_info = get_document_info_from_dxf(dxf_resultado, log=None)

        if not isinstance(res_info, tuple):
            log.write(f"[JL][ERRO] get_document_info_from_dxf retornou {type(res_info).__name__}: {res_info!r}")
            return False

        if len(res_info) == 6:
            doc, lines, arcs, perimeter_dxf, area_dxf, _boundary_points = res_info
        elif len(res_info) == 5:
            doc, lines, arcs, perimeter_dxf, area_dxf = res_info
            _boundary_points = None
        else:
            log.write(f"[JL][ERRO] get_document_info_from_dxf retornou tupla de tamanho inesperado: len={len(res_info)}")
            return False

        if not doc or not lines:
            log.write("❌ Documento DXF inválido ou sem linhas.")
            return False

        msp = doc.modelspace()
        v1 = lines[0][0]
        distance_az_v1 = calculate_distance(ponto_az, v1)
        azimute_az_v1  = calculate_azimuth(ponto_az, v1)
        # Para o seu create_memorial_document do JL
        distance = distance_az_v1
        azimuth  = azimute_az_v1

        # === chama a create_memorial_descritivo com a MESMA assinatura do DECOPA ===
        # Observação: usamos tipo="JL" só para manter o parâmetro.
        excel_output = create_memorial_descritivo(
            doc=doc,
            msp=msp,
            lines=lines,
            arcs=arcs,
            proprietario=proprietario,
            matricula=matricula,
            caminho_salvar=caminho_salvar,       # CONCLUIDO
            excel_file_path=excel_path,          # sua planilha de uma aba ("Confrontantes")
            ponto_az=ponto_az,                   # mesmo não sendo usado, mantemos a assinatura
            distance_az_v1=distance_az_v1,
            azimute_az_v1=azimute_az_v1,
            ponto_inicial_real=ponto_inicial_real,
            tipo="JL",
            uuid_prefix=uuid_prefix,
            sentido_poligonal=sentido_poligonal,
        )

        if not excel_output:
            log.write("[ERRO] Falha ao gerar memorial descritivo (XLSX/DXF).")
            return False

        # === DOCX no padrão JL (mantendo seus parâmetros existentes) ===
        template_path = os.path.join("templates_doc", "MODELO_TEMPLATE_DOC_JL_CORRETO.docx")
        docx_path = os.path.join(caminho_salvar, f"Memorial_MAT_{safe_mat}.docx")
        try:
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
                log=log,  # seu JL aceita log
            )
            log.write(f"[JL] DOCX salvo em: {docx_path}")
        except Exception as e:
            log.write(f"[JL] DOCX opcional não gerado: {e}")

        log.write("✅ Processamento finalizado com sucesso.")
        return True

    except Exception as e:
        tb = traceback.format_exc()
        try:
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(f"\n❌ ERRO inesperado em executar_memorial_jl:\n{tb}\n")
        except Exception:
            pass
        logger.exception("Erro inesperado em executar_memorial_jl")
        return False


