# executaveis_avaliacao/utils_json.py
# -*- coding: utf-8 -*-

from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd

# ───────────────────────────── paths ─────────────────────────────
BASE_DIR: Path = Path(__file__).resolve().parents[1]
TMP_DIR: Path = BASE_DIR / "static" / "tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

# ───────────────────────── chaves padrão ─────────────────────────
# Campos obrigatórios e fatores por amostra que queremos SEMPRE garantir no JSON
ESSENTIAL_KEYS = [
    "idx", "valor_total", "area", "valor_unitario", "cidade", "fonte",
    "LATITUDE", "LONGITUDE", "ativo"
]
FACTOR_KEYS = [
    "APROVEITAMENTO",
    "BOA TOPOGRAFIA?",
    "PEDOLOGIA ALAGÁVEL? ",   # (observa o espaço no final conforme planilhas existentes)
    " ESQUINA?",              # (observa o espaço inicial conforme planilhas existentes)
    "PAVIMENTACAO?",
    "ACESSIBILIDADE?",
    "DISTANCIA CENTRO",
]

# Mapeamento de aliases → chave canônica (o que fica no JSON)
ALIASES = {
    # valores/área/idx
    "VALOR TOTAL": "valor_total",
    "AREA TOTAL": "area",
    "AM": "idx",
    # distancia
    "distancia_centro": "DISTANCIA CENTRO",
    # pav/acess topo/pedo/esquina
    "PAVIMENTAÇÃO?": "PAVIMENTACAO?",
    "PAVIMENTACAO ?": "PAVIMENTACAO?",
    "ACESSIBILIDADE ?": "ACESSIBILIDADE?",
    "BOA TOPOGRAFIA ?": "BOA TOPOGRAFIA?",
    "PEDOLOGIA ALAGAVEL?": "PEDOLOGIA ALAGÁVEL? ",
    "ESQUINA?": " ESQUINA?",
}

# ─────────────────────── helpers internos ────────────────────────
def _coerce_float(x: Any, default: float = 0.0) -> float:
    try:
        if isinstance(x, str):
            x = x.replace("R$", "").replace(".", "").replace(",", ".").strip()
        v = float(x)
        if v != v:  # NaN
            return default
        return v
    except Exception:
        return default


def _alias_to_canonical(key: str) -> str:
    return ALIASES.get(key, key)


def _row_to_dict(row: Any) -> Dict[str, Any]:
    # aceita Series/Mapping
    if hasattr(row, "to_dict"):
        d = row.to_dict()
    else:
        d = dict(row or {})
    # normaliza aliases de primeira camada
    out = {}
    for k, v in d.items():
        out[_alias_to_canonical(k)] = v
    return out


def _normalize_amostras(
    amostras: Union[pd.DataFrame, List[Dict[str, Any]]]
) -> List[Dict[str, Any]]:
    """
    Normaliza amostras para o formato do snapshot v2 e
    GARANTE a presença de todos os fatores em FACTOR_KEYS (mesmo None).
    Mantém quaisquer campos extras enviados, mas com nomes canônicos.
    """
    registros: List[Dict[str, Any]] = []

    # Converte DataFrame para lista de dicts canônicos
    if isinstance(amostras, pd.DataFrame):
        # renomeia aliases de colunas para canônicas
        df = amostras.copy()
        rename_map = {src: dst for src, dst in ALIASES.items() if src in df.columns}
        if rename_map:
            df.rename(columns=rename_map, inplace=True)
        it = (df.itertuples(index=False, name=None), df.columns.tolist())
        tuples, cols = it
        for tup in tuples:
            row_map = dict(zip(cols, tup))
            registros.append(_normalize_um(row_map))
        return registros

    # Lista de dicts
    for a in amostras:
        registros.append(_normalize_um(a))

    return registros


def _normalize_um(a: Dict[str, Any]) -> Dict[str, Any]:
    """Normaliza um único registro de amostra, garantindo ESSENTIAL_KEYS e FACTOR_KEYS."""
    d = _row_to_dict(a)

    # Campos essenciais com defaults seguros
    idx = int(_coerce_float(d.get("idx", d.get("AM", 0)), 0))
    vt = _coerce_float(d.get("valor_total", d.get("VALOR TOTAL", 0.0)), 0.0)
    ar = _coerce_float(d.get("area", d.get("AREA TOTAL", 0.0)), 0.0)

    # valor_unitario (recalcula se não existir ou inválido)
    vu = d.get("valor_unitario", None)
    if vu is None:
        vu = (vt / ar) if ar > 0 else 0.0
    else:
        vu = _coerce_float(vu, (vt / ar) if ar > 0 else 0.0)

    base = {
        "idx": idx,
        "valor_total": vt,
        "area": ar,
        "valor_unitario": vu,
        "cidade": str(d.get("cidade", "") or ""),
        "fonte": str(d.get("fonte", "") or ""),
        "LATITUDE": (_coerce_float(d.get("LATITUDE", 0.0), 0.0)
                     if d.get("LATITUDE", None) is not None else None),
        "LONGITUDE": (_coerce_float(d.get("LONGITUDE", 0.0), 0.0)
                      if d.get("LONGITUDE", None) is not None else None),
        "ativo": bool(d.get("ativo", True)),
    }

    # Garante todos os fatores (se não existirem, ficam None ou 0.0 no caso de distância)
    factors = {
        "APROVEITAMENTO": d.get("APROVEITAMENTO", None),
        "BOA TOPOGRAFIA?": d.get("BOA TOPOGRAFIA?", d.get("BOA TOPOGRAFIA ?", None)),
        "PEDOLOGIA ALAGÁVEL? ": d.get("PEDOLOGIA ALAGÁVEL? ", d.get("PEDOLOGIA ALAGAVEL?", None)),
        " ESQUINA?": d.get(" ESQUINA?", d.get("ESQUINA?", None)),
        "PAVIMENTACAO?": d.get("PAVIMENTACAO?", d.get("PAVIMENTAÇÃO?", d.get("PAVIMENTACAO ?", None))),
        "ACESSIBILIDADE?": d.get("ACESSIBILIDADE?", d.get("ACESSIBILIDADE ?", None)),
        "DISTANCIA CENTRO": _coerce_float(d.get("DISTANCIA CENTRO", d.get("distancia_centro", 0.0)), 0.0),
    }

    # Junta: preserva quaisquer campos extras (sem sobrescrever os normalizados)
    extras = {k: v for k, v in d.items() if k not in {**base, **factors}}
    return {**extras, **base, **factors}


# ───────────────────── funções públicas v2 ──────────────────────
def salvar_entrada_corrente_json(
    uuid_execucao: str,
    dados_avaliando: Dict[str, Any],
    fatores_do_usuario: Dict[str, Any],
    dataframe_amostras: Union[pd.DataFrame, List[Dict[str, Any]]],
    *,
    variavel_dependente: str,
    variaveis_independentes: List[str],
    transformacoes_por_variavel: Dict[str, Any],
    restricoes_de_sinal: Dict[str, Any],
    amostras_desabilitadas: List[int],
    parametros_modelo: Dict[str, Any],
    arquivos: Optional[Dict[str, Any]] = None,
    app_version: str = "AVALIACAO-2025.08",
    random_seed: Optional[int] = None,
) -> str:
    """
    Salva o snapshot COMPLETO (schema v2). As amostras passam por normalização
    e já saem com TODOS os fatores garantidos.
    """
    amostras_norm = _normalize_amostras(dataframe_amostras)

    payload: Dict[str, Any] = {
        "meta": {
            "schema_ver": 2,
            "uuid_execucao": uuid_execucao,
            "app_version": app_version,
        },
        "dados_avaliando": dict(dados_avaliando or {}),
        "amostras": amostras_norm,
        "fatores_do_usuario": dict(fatores_do_usuario or {}),
        "config_modelo": {
            "variavel_dependente": variavel_dependente,
            "variaveis_independentes": list(variaveis_independentes or []),
            "transformacoes_por_variavel": dict(transformacoes_por_variavel or {}),
            "restricoes_de_sinal": dict(restricoes_de_sinal or {}),
            "amostras_desabilitadas": list(amostras_desabilitadas or []),
            "parametros": dict(parametros_modelo or {}),
            "random_seed": random_seed,
        },
        "arquivos": dict(arquivos or {}),
        "resultados_parciais": {},
    }

    caminho_arquivo = TMP_DIR / f"{uuid_execucao}_entrada_corrente.json"
    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return str(caminho_arquivo)


def carregar_entrada_corrente_json(uuid_execucao: str) -> Dict[str, Any]:
    """
    Carrega o snapshot. Se encontrar schema antigo (v1), migra em memória para v2
    e também garante os fatores nas amostras migradas.
    """
    caminho_arquivo = TMP_DIR / f"{uuid_execucao}_entrada_corrente.json"
    if not caminho_arquivo.exists():
        raise FileNotFoundError(f"JSON não encontrado: {caminho_arquivo}")

    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        data = json.load(f)

    schema_ver = (data.get("meta") or {}).get("schema_ver")
    if schema_ver == 2:
        # Mesmo no v2, podemos garantir que as amostras tenham os fatores padrão
        data["amostras"] = _normalize_amostras(data.get("amostras", []))
        return data

    # ── MIGRAÇÃO simples v1 → v2 ──
    dados_avaliando = data.get("dados_avaliando", {})
    fatores_do_usuario = data.get("fatores_do_usuario", {})
    amostras = data.get("amostras", [])

    amostras_norm = _normalize_amostras(amostras)

    migrated: Dict[str, Any] = {
        "meta": {
            "schema_ver": 2,
            "uuid_execucao": data.get("uuid_execucao")
            or (data.get("meta") or {}).get("uuid_execucao"),
            "app_version": (data.get("meta") or {}).get("app_version", "AVALIACAO-2025.08"),
        },
        "dados_avaliando": dados_avaliando,
        "amostras": amostras_norm,
        "fatores_do_usuario": fatores_do_usuario,
        "config_modelo": {
            "variavel_dependente": data.get("variavel_dependente") or "VALOR UNITARIO",
            "variaveis_independentes": data.get("variaveis_independentes") or [],
            "transformacoes_por_variavel": data.get("transformacoes_por_variavel") or {},
            "restricoes_de_sinal": data.get("restricoes_de_sinal") or {},
            "amostras_desabilitadas": data.get("amostras_desabilitadas") or [],
            "parametros": data.get("parametros_modelo") or {},
            "random_seed": data.get("random_seed"),
        },
        "arquivos": data.get("arquivos") or {
            "fotos_imovel": data.get("fotos_imovel") or [],
            "fotos_adicionais": data.get("fotos_adicionais") or [],
            "fotos_proprietario": data.get("fotos_proprietario") or [],
            "fotos_planta": data.get("fotos_planta") or [],
        },
        "resultados_parciais": data.get("resultados_parciais") or {},
    }
    return migrated


# ─────────────── wrapper legacy (compatibilidade) ───────────────
def salvar_entrada_corrente_json_legacy(
    dados_imovel: Dict[str, Any],
    fatores_usuario: Dict[str, Any],
    amostras: List[Dict[str, Any]],
    uuid_execucao: str,
    fotos_imovel: Optional[List] = None,
    fotos_adicionais: Optional[List] = None,
    fotos_proprietario: Optional[List] = None,
    fotos_planta: Optional[List] = None,
) -> str:
    """
    Compat com chamadas antigas:
      salvar_entrada_corrente_json(dados_imovel, fatores_usuario, amostras, uuid, ...)
    Redireciona para o v2 com defaults e normalização completa.
    """
    arquivos = {
        "fotos_imovel": fotos_imovel or [],
        "fotos_adicionais": fotos_adicionais or [],
        "fotos_proprietario": fotos_proprietario or [],
        "fotos_planta": fotos_planta or [],
    }
    return salvar_entrada_corrente_json(
        uuid_execucao=uuid_execucao,
        dados_avaliando=dados_imovel,
        fatores_do_usuario=fatores_usuario,
        dataframe_amostras=amostras,  # lista de dicts aceita e normalizada
        variavel_dependente="VALOR UNITARIO",
        variaveis_independentes=[],
        transformacoes_por_variavel={},
        restricoes_de_sinal={},
        amostras_desabilitadas=[],
        parametros_modelo={},
        arquivos=arquivos,
        app_version="AVALIACAO-2025.08",
        random_seed=None,
    )
