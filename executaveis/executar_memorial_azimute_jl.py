# --- IMPORTS DE MÓDULO (fora da função; necessários para helpers e cálculos) ---
import os
import math
import traceback
import re
from pathlib import Path
import time
from glob import glob
import shutil


# fallback simples caso sanitize_filename não exista no módulo
def _sanitize_filename(s):
    return re.sub(r'[^0-9A-Za-z._-]+', '_', str(s)).strip('_')

try:
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
    import logging, shutil
    from .memoriais_JL import (
        limpar_dxf_e_inserir_ponto_az,   # ← EXATAMENTE como no DECOPA
        get_document_info_from_dxf,
        create_memorial_descritivo,
        create_memorial_document,
    )

    # ===== Preparação de pastas =====
    Path(caminho_salvar).mkdir(parents=True, exist_ok=True)
    Path(Path(log_path).parent).mkdir(parents=True, exist_ok=True)

    # ===== Logging no console (Render) se ainda não houver handler =====
    root = logging.getLogger()
    if not root.handlers:
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s [%(name)s] %(message)s"
        )

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    # ===== Cabeçalho do arquivo de log =====
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write("🟢 LOG JL iniciado\n")
            fh.write(f"[JL] sentido_poligonal={sentido_poligonal}\n")
    except Exception:
        logger.warning(f"[JL] Falha ao iniciar o log: {log_path}")

    # Writer simples p/ arquivo + espelho no console
    def _log_file(msg: str):
        try:
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(msg if msg.endswith("\n") else msg + "\n")
        except Exception:
            logger.warning(f"[JL] Falha ao escrever no log: {log_path}")
        try:
            logger.info(str(msg).strip())
        except Exception:
            pass

    class _LogWriter:
        def __init__(self, file_path):
            self.file_path = file_path
        def write(self, msg):
            _log_file(str(msg))

    log = _LogWriter(log_path)
    _log_file(f"[JL] log_path: {log_path}")
    _log_file(f"[JL] Iniciando executar_memorial_jl | sentido_poligonal={sentido_poligonal}")

    try:
        # === UUID prefix (basename do pai da pasta CONCLUIDO) ===
        uuid_prefix = os.path.basename(os.path.dirname(os.path.normpath(caminho_salvar))) or "JL"
        safe_mat = _sanitize_filename(matricula)

        # === cria caminho do DXF limpo ===
        nome_limpo_dxf = f"{uuid_prefix}_JL_{safe_mat}_LIMPO.dxf"
        caminho_dxf_limpo = os.path.join(caminho_salvar, nome_limpo_dxf)

        # === limpeza do DXF — padrão DECOPA ===
        _log_file("[JL] Limpando DXF com: limpar_dxf_e_inserir_ponto_az")
        try:
            res = limpar_dxf_e_inserir_ponto_az(dxf_path, caminho_dxf_limpo)
        except Exception:
            tb = traceback.format_exc()
            _log_file("❌ EXCEÇÃO em limpar_dxf_e_inserir_ponto_az:")
            _log_file(tb)
            return log_path, []   # <<< RETORNO PADRÃO (log, lista vazia)

        _log_file(f"[JL] retorno limpar_dxf_e_inserir_ponto_az: tipo={type(res).__name__} valor={repr(res)[:300]}")
        if not isinstance(res, tuple) or len(res) != 3:
            _log_file(f"[ERRO] Esperava tupla(dxf_resultado, ponto_az, ponto_inicial); recebi {type(res).__name__}")
            return log_path, []

        dxf_resultado, ponto_az, ponto_inicial = res
        ponto_inicial_real = ponto_inicial  # compatibilidade
        if not ponto_az:
            ponto_az = (0.0, 0.0)  # neutro só para manter assinatura

        # === leitura das entidades do DXF ===
        try:
            res_info = get_document_info_from_dxf(dxf_resultado)
        except TypeError:
            res_info = get_document_info_from_dxf(dxf_resultado, log=None)
        except Exception:
            tb = traceback.format_exc()
            _log_file("❌ EXCEÇÃO em get_document_info_from_dxf:")
            _log_file(tb)
            return log_path, []

        _log_file(f"[JL] retorno get_document_info_from_dxf: tipo={type(res_info).__name__} valor={repr(res_info)[:300]}")
        if not isinstance(res_info, tuple):
            _log_file(f"[ERRO] Esperava tupla(doc, lines, arcs, perimeter_dxf, area_dxf[, boundary]); recebi {type(res_info).__name__}")
            return log_path, []

        if len(res_info) == 6:
            doc, lines, arcs, perimeter_dxf, area_dxf, _boundary_points = res_info
        elif len(res_info) == 5:
            doc, lines, arcs, perimeter_dxf, area_dxf = res_info
            _boundary_points = None
        else:
            _log_file(f"[ERRO] Tamanho inesperado da tupla de retorno: len={len(res_info)}")
            return log_path, []

        if not doc or not lines:
            _log_file("❌ Documento DXF inválido ou sem entidades de linha.")
            return log_path, []

        msp = doc.modelspace()
        v1 = lines[0][0]
        distance_az_v1 = calculate_distance(ponto_az, v1)
        azimute_az_v1  = calculate_azimuth(ponto_az, v1)
        distance = distance_az_v1     # para o DOCX JL
        azimuth  = azimute_az_v1      # para o DOCX JL

        
                # === Marco temporal para identificar arquivos salvos após a create ===
        t0 = time.time()

        # === create_memorial_descritivo — MESMA assinatura do DECOPA ===
        excel_output = create_memorial_descritivo(
            doc=doc,
            msp=msp,
            lines=lines,
            arcs=arcs,
            proprietario=proprietario,
            matricula=matricula,
            caminho_salvar=caminho_salvar,       # CONCLUIDO
            excel_file_path=excel_path,          # Excel com aba "Confrontantes"
            ponto_az=ponto_az,
            distance_az_v1=distance_az_v1,
            azimute_az_v1=azimute_az_v1,
            ponto_inicial_real=ponto_inicial_real,
            tipo="JL",
            uuid_prefix=uuid_prefix,
            sentido_poligonal=sentido_poligonal,
        )

        if not excel_output:
            _log_file("[ERRO] Falha ao gerar memorial descritivo (XLSX/DXF).")
            return log_path, []

        # === DOCX no padrão JL (define docx_path ANTES de montar a lista de arquivos) ===
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
                log=_LogWriter(log_path),  # aceita .write(...)
            )
            _log_file(f"[JL] DOCX salvo em: {docx_path}")
        except Exception as e:
            _log_file(f"[JL] DOCX opcional não gerado: {e}")
            docx_path = None  # segue sem DOCX se falhar

        # === Seleciona o DXF ANOTADO e consolida o final ===
        dxfs = glob(os.path.join(caminho_salvar, "*.dxf"))
        _log_file(f"[JL] DXFs no CONCLUIDO: {len(dxfs)} -> {[os.path.basename(p) for p in dxfs]}")

        # Preferir DXFs salvos/alterados após a create (se t0 existir)
        if dxfs:
            try:
                base_list = [p for p in dxfs if os.path.getmtime(p) >= t0 - 0.5] if 't0' in locals() else dxfs
                base_list = base_list or dxfs  # fallback se filtro ficar vazio
                annotated_dxf = max(base_list, key=lambda p: os.path.getmtime(p))
            except Exception as e:
                _log_file(f"[JL] Aviso: falha ao escolher DXF anotado por mtime: {e}")
                annotated_dxf = dxfs[-1]  # fallback
        else:
            annotated_dxf = dxf_resultado if os.path.exists(dxf_resultado) else None

        # Remove o LIMPO só se NÃO for o anotado
        if annotated_dxf and os.path.exists(caminho_dxf_limpo) and \
           os.path.abspath(annotated_dxf) != os.path.abspath(caminho_dxf_limpo):
            try:
                os.remove(caminho_dxf_limpo)
                _log_file(f"[JL] DXF LIMPO removido: {caminho_dxf_limpo}")
            except Exception as e:
                _log_file(f"[JL] Aviso: não foi possível remover DXF LIMPO: {e}")

        # Copia para nome final
        final_dxf_path = None
        if annotated_dxf and os.path.exists(annotated_dxf):
            final_dxf_path = os.path.join(caminho_salvar, f"Memorial_{safe_mat}.dxf")
            try:
                if os.path.abspath(annotated_dxf) != os.path.abspath(final_dxf_path):
                    shutil.copyfile(annotated_dxf, final_dxf_path)
                _log_file(f"[JL] DXF final (ANOTADO): {final_dxf_path} (origem: {annotated_dxf})")
            except Exception as e:
                _log_file(f"[JL] Aviso: cópia do DXF anotado falhou ({e}); usando {annotated_dxf}")
                final_dxf_path = annotated_dxf

        # === Retorno no formato que a rota espera ===
        arquivos = [p for p in [excel_output, final_dxf_path, docx_path] if p and os.path.exists(p)]
        if not arquivos:
            _log_file("[ERRO] Nenhum arquivo encontrado para retorno.")
            return log_path, []

        _log_file("✅ Processamento finalizado com sucesso.")
        return log_path, arquivos


       
    except Exception:
        tb = traceback.format_exc()
        _log_file("❌ ERRO inesperado em executar_memorial_jl:")
        _log_file(tb)
        logger.exception("Erro inesperado em executar_memorial_jl")
        # <<< RETORNO NO PADRÃO ANTIGO >>>
        return log_path, []
