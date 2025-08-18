import os, sys, logging, json, shutil, uuid, glob
from datetime import datetime
from pathlib import Path

# ========= RUN CONTEXT =========
RUN_UUID = os.environ.get("RUN_UUID") or uuid.uuid4().hex[:8]
os.environ["RUN_UUID"] = RUN_UUID
RUN_BASE = Path(os.environ.get("RUN_BASE", "/opt/render/project/src/tmp"))
RUN_DIR = RUN_BASE / RUN_UUID
CONCLUIDO_DIR = RUN_DIR / "CONCLUIDO"
CONCLUIDO_DIR.mkdir(parents=True, exist_ok=True)

# Log único desta execução: /tmp/<uuid>/CONCLUIDO/exec_<uuid>.log
LOG_FILE = CONCLUIDO_DIR / f"exec_{RUN_UUID}.log"
logger = logging.getLogger("memorial")
logger.setLevel(logging.INFO)
if not any(isinstance(h, logging.FileHandler) and getattr(h, "_run_uuid", None) == RUN_UUID for h in logger.handlers):
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh._run_uuid = RUN_UUID
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(sh)
# ===============================

from preparar_arquivos import preparar_arquivos
from poligonal_fechada import main_poligonal_fechada
from compactar_arquivos import main_compactar_arquivos

BASE_DIR = Path(__file__).resolve().parents[1]
CAMINHO_PUBLICO = BASE_DIR / 'static' / 'arquivos'
CAMINHO_PUBLICO.mkdir(parents=True, exist_ok=True)

def main():
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print("Uso: python main.py <cidade> <caminho_excel> <caminho_dxf> [sentido_poligonal]")
        sys.exit(1)

    cidade = sys.argv[1]
    uuid_str = RUN_UUID  # 🔴 use o MESMO UUID da execução em TODO o pipeline
    cidade_formatada = cidade.replace(" ", "_")
    caminho_excel = sys.argv[2]
    caminho_dxf   = sys.argv[3]
    sentido_poligonal = sys.argv[4] if len(sys.argv) == 5 else 'horario'
    logger.info(f"RUN_UUID={RUN_UUID} | Sentido={sentido_poligonal}")

    caminho_template = str(BASE_DIR / "templates_doc" / "Memorial_modelo_padrao.docx")
    if not os.path.exists(caminho_template):
        logger.error(f"Template não encontrado em '{caminho_template}'.")
        _write_manifest(success=False, zip_files=[])
        sys.exit(1)

    variaveis = preparar_arquivos(cidade, caminho_excel, caminho_dxf, str(BASE_DIR), uuid_str)
    if not variaveis:
        logger.error("Erro ao preparar arquivos. Encerrando execução.")
        _write_manifest(success=False, zip_files=[])
        sys.exit(1)

    logger.info("✅ Preparação dos arquivos concluída.")

    main_poligonal_fechada(
        uuid_str,
        variaveis["arquivo_excel_recebido"],
        variaveis["arquivo_dxf_recebido"],
        variaveis["diretorio_preparado"],
        variaveis["diretorio_concluido"],
        caminho_template,
        sentido_poligonal
    )
    logger.info("✅ Processamento da poligonal fechada concluído.")

    # Sinalização: existe algo para zipar?
    tem_algum = any(glob.glob(os.path.join(variaveis["diretorio_concluido"], f"{uuid_str}_FECHADA_{t}_*.xlsx"))
                    for t in ("ETE","REM","SER","ACE"))
    if not tem_algum:
        logger.error("❌ Nenhum XLSX FECHADA gerado em %s. Compactação pode não gerar ZIP.",
                     variaveis["diretorio_concluido"])

    main_compactar_arquivos(variaveis["diretorio_concluido"], cidade_formatada, uuid_str)
    logger.info("✅ Compactação concluída.")

    # (Opcional) Copiar ZIPs para static/arquivos — o front novo não precisa disso, mas mantive
    try:
        zips = list(Path(variaveis["diretorio_concluido"]).glob("*.zip"))
        for p in zips:
            shutil.copy2(p, CAMINHO_PUBLICO / p.name)
            logger.info(f"📦 ZIP copiado: {p.name}")
        if not zips:
            logger.warning("⚠️ Nenhum ZIP encontrado para copiar.")
    except Exception as e:
        logger.error(f"❌ Erro ao copiar ZIPs: {e}")

    # Manifesto da execução (AGORA sim, depois de tudo pronto)
    zip_files = [p.name for p in Path(variaveis["diretorio_concluido"]).glob("*.zip")]
    _write_manifest(success=bool(zip_files), zip_files=zip_files)

def _write_manifest(success: bool, zip_files):
    try:
        data = {
            "uuid": RUN_UUID,
            "success": bool(success),
            "log": LOG_FILE.name,
            "zip_files": list(zip_files),
            "finished_at": datetime.now().isoformat(timespec="seconds"),
        }
        (CONCLUIDO_DIR / "RUN.json").write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                                encoding="utf-8")
        logger.info(f"[RUN] Manifesto salvo: {CONCLUIDO_DIR/'RUN.json'}")
    finally:
        # garante que o log está escrito em disco antes do Flask servir o download
        for h in list(logging.getLogger("memorial").handlers):
            try:
                h.flush()
            except Exception:
                pass

if __name__ == "__main__":
    main()
