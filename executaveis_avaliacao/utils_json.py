# executaveis_avaliacao/utils_json.py

import json
import os

def salvar_entrada_corrente_json(
    dados_imovel,
    fatores_usuario,
    amostras,
    uuid_execucao,
    fotos_imovel=None,
    fotos_adicionais=None,
    fotos_proprietario=None,
    fotos_planta=None
):
    """
    Gera e salva o arquivo entrada_corrente.json contendo:
    - Dados do imóvel
    - Fatores selecionados
    - Lista de amostras com campo "ativo": True
    - Caminhos das fotos (opcionais)
    """
    estrutura = {
        "dados_avaliando": dados_imovel,
        "fatores_do_usuario": fatores_usuario,
        "amostras": [],
        "fotos_imovel": fotos_imovel or [],
        "fotos_adicionais": fotos_adicionais or [],
        "fotos_proprietario": fotos_proprietario or [],
        "fotos_planta": fotos_planta or []
    }

    for a in amostras:
        estrutura["amostras"].append({
            "idx": a.get("AM") or a.get("idx"),
            "valor_total": float(a.get("VALOR TOTAL", 0)),
            "area": float(a.get("AREA TOTAL", 0)),
            "valor_unitario": float(a.get("VALOR UNITARIO", 0)),
            "cidade": a.get("CIDADE", ""),
            "fonte": a.get("FONTE", ""),
            "coordenada_x": a.get("X", None),
            "coordenada_y": a.get("Y", None),
            "ativo": True
        })


    pasta_saida = f"static/tmp"
    os.makedirs(pasta_saida, exist_ok=True)

    caminho_arquivo = os.path.join(pasta_saida, f"{uuid_execucao}_entrada_corrente.json")

    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(estrutura, f, indent=2, ensure_ascii=False)

    return caminho_arquivo
