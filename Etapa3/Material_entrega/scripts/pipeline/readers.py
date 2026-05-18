from __future__ import annotations

import pandas as pd

from pipeline.config import CONFIG, PATHS
from pipeline.utils import (
    extrair_ano,
    listar_excel,
    localizar_coluna,
    renomear_colunas_normalizadas,
)


SICONFI_COLUMNS = {
    "instituicao": "instituicao",
    "cod.ibge": "cod_ibge",
    "uf": "uf",
    "populacao": "populacao",
    "coluna": "coluna",
    "conta": "conta",
    "identificador da conta": "identificador_conta",
    "valor": "valor",
}


def ler_siconfi(caminho) -> pd.DataFrame:
    df = pd.read_excel(caminho, skiprows=3)
    df = renomear_colunas_normalizadas(df, SICONFI_COLUMNS)
    df["ano"] = extrair_ano(caminho)
    df["cod_ibge"] = pd.to_numeric(df["cod_ibge"], errors="coerce").astype("Int64")
    df["populacao"] = pd.to_numeric(df["populacao"], errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return df[df["uf"].eq(CONFIG.uf)].copy()


def ler_receitas_siconfi() -> pd.DataFrame:
    frames = [ler_siconfi(caminho) for caminho in listar_excel(PATHS.receitas, (".xlsx",))]
    return pd.concat(frames, ignore_index=True)


def ler_despesas_siconfi() -> pd.DataFrame:
    frames = [ler_siconfi(caminho) for caminho in listar_excel(PATHS.despesas, (".xlsx",))]
    return pd.concat(frames, ignore_index=True)


def ler_pib_municipal() -> pd.DataFrame:
    frames = []
    for caminho in listar_excel(PATHS.pib, (".xlsx",)):
        df = pd.read_excel(caminho, sheet_name=0)
        colunas = [
            localizar_coluna(df, ["ano"]),
            localizar_coluna(df, ["sigla", "unidade", "federacao"]),
            localizar_coluna(df, ["codigo", "municipio"]),
            localizar_coluna(df, ["nome", "municipio"]),
            localizar_coluna(df, ["produto interno bruto", "precos correntes", "1.000"]),
            localizar_coluna(df, ["produto interno bruto per capita"]),
        ]
        parte = df[colunas].copy()
        parte.columns = [
            "ano",
            "uf",
            "cod_ibge",
            "municipio_pib",
            "pib_mil_reais",
            "pib_per_capita_ibge",
        ]
        frames.append(parte)

    pib = pd.concat(frames, ignore_index=True)
    pib["ano"] = pd.to_numeric(pib["ano"], errors="coerce").astype("Int64")
    pib["cod_ibge"] = pd.to_numeric(pib["cod_ibge"], errors="coerce").astype("Int64")
    pib["pib_mil_reais"] = pd.to_numeric(pib["pib_mil_reais"], errors="coerce")
    pib["pib_per_capita_ibge"] = pd.to_numeric(
        pib["pib_per_capita_ibge"], errors="coerce"
    )
    pib = pib[pib["uf"].eq(CONFIG.uf)].copy()
    pib["pib"] = pib["pib_mil_reais"] * 1000
    pib = pib.sort_values(["cod_ibge", "ano"])
    pib["crescimento_pib"] = pib.groupby("cod_ibge")["pib"].pct_change()
    return pib[
        [
            "ano",
            "cod_ibge",
            "municipio_pib",
            "pib",
            "pib_per_capita_ibge",
            "crescimento_pib",
        ]
    ]
