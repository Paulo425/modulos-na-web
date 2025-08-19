# executaveis/exec_ctx.py
import os, sys, re, json, logging
from pathlib import Path
from datetime import datetime

# -------------------------------
# 1) ID_EXECUCAO: única fonte de verdade
# -------------------------------
def _get_id_execucao():
    v = os.environ.get("ID_EXECUCAO")
    if not v:
        args = sys.argv[1:]
        # --id-execucao=XXXX
        for a in args:
            if a.startswith("--id-execucao="):
                v = a.split("=", 1)[1].strip()
                break
        # --id-execucao XXXX
        if not v:
            for i, a in enumerate(args):
                if a == "--id-execucao" and i + 1 < len(args):
                    v = args[i + 1].strip()
                    break
        # fallback de compatibilidade, se alguém ainda usa --run-uuid
        if not v:
            for a in args:
                if a.startswith("--run-uuid="):
                    v = a.split("=", 1)[1].strip()
                    break

    if not v:
        sys.stderr.write("[ERRO] ID_EXECUCAO não informado (env ID_EXECUCAO ou --id-execucao ...).\n")
        sys.exit(2)
    if not re.fullmatch(r"[A-Za-z0-9_\-]{6,64}", v):
        sys.stderr.write(f"[ERRO] ID_EXECUCAO inválido: {v}\n")
        sys.exit(2)
    return v

ID_EXECUCAO = _get_id_execucao()

# -------------------------------
# 2) BASE_DIR e estrutura de diretórios
#     -> use env BASE_DIR para sobrepor, se necessário
# -------------------------------
BASE_DIR = os.environ.get("BASE_DIR") or str(Path(__file__).resolve().parent)
DIR_RUN  = os.path.join(BASE_DIR, "tmp", ID_EXECUCAO)
DIR_REC  = os.path.join(DIR_RUN, "RECEBIDO")
DIR_PREP = os.path.join(DIR_RUN, "PREPARADO")
DIR_CONC = os.path.join(DIR_RUN, "CONCLUIDO")

def ensure_dirs():
    for d in (DIR_RUN, DIR_REC, DIR_PREP, DIR_CONC):
        os.makedirs(d, exist_ok=True)

# -------------------------------
# 3) Metadata da execução (anti-mistura)
# -------------------------------
def write_metadata_if_missing(extra: dict | None = None):
    ensure_dirs()
    meta = {
        "id_execucao": ID_EXECUCAO,
        "criado_em": datetime.utcnow().isoformat() + "Z",
        "base_dir": BASE_DIR,
        "dir_run": DIR_RUN,
    }
    if extra:
        meta.update(extra)
    path = os.path.join(DIR_RUN, "exec_metadata.json")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    return path

def validate_metadata():
    path = os.path.join(DIR_RUN, "exec_metadata.json")
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        meta = json.load(f)
    if meta.get("id_execucao") != ID_EXECUCAO:
        sys.stderr.write("[ERRO] ID_EXECUCAO do processo não confere com exec_metadata.json.\n")
        sys.exit(3)

# -------------------------------
# 4) Logging unificado (stdout + arquivo)
# -------------------------------
def log_path():
    return os.path.join(DIR_CONC, f"exec_{ID_EXECUCAO}.log")

def setup_logger(name="pipeline", level=logging.INFO):
    ensure_dirs()
    validate_metadata()
    lp = log_path()

    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    # arquivo
    fh = logging.FileHandler(lp, encoding="utf-8")
    fh.setFormatter(fmt)
    fh.setLevel(level)
    logger.addHandler(fh)

    # stdout
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    sh.setLevel(level)
    logger.addHandler(sh)

    logger.info("ID_EXECUCAO=%s", ID_EXECUCAO)
    logger.info("DIR_RUN=%s", DIR_RUN)
    logger.info("Pastas: REC=%s | PREP=%s | CONC=%s", DIR_REC, DIR_PREP, DIR_CONC)
    return logger

# inicialização mínima quando importado
ensure_dirs()
write_metadata_if_missing()
