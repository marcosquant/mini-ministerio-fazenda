from __future__ import annotations

import re

import numpy as np
import pandas as pd

from pipeline.config import COLUNAS_BASE_FINAL, CONFIG


def nome_municipio(instituicao: str) -> str:
    nome = str(instituicao)
    nome = re.sub(r"^Prefeitura Municipal de\s+", "", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\s+-\s+[A-Z]{2}$", "", nome)
    return nome.strip()


def preparar_populacao(receitas: pd.DataFrame, despesas: pd.DataFrame) -> pd.DataFrame:
    pop_receitas = (
        receitas[["ano", "cod_ibge", "uf", "instituicao", "populacao"]]
        .drop_duplicates()
        .rename(columns={"populacao": "populacao_receitas"})
    )
    pop_despesas = (
        despesas[["ano", "cod_ibge", "populacao"]]
        .drop_duplicates()
        .rename(columns={"populacao": "populacao_despesas"})
    )

    pop = pop_receitas.merge(pop_despesas, on=["ano", "cod_ibge"], how="outer")
    divergencias = pop[
        pop["populacao_receitas"].notna()
        & pop["populacao_despesas"].notna()
        & pop["populacao_receitas"].ne(pop["populacao_despesas"])
    ]
    if not divergencias.empty:
        raise ValueError(
            "Foram encontradas divergencias entre populacao de receitas e despesas."
        )

    pop["populacao"] = pop["populacao_receitas"].combine_first(pop["populacao_despesas"])
    pop["municipio"] = pop["instituicao"].map(nome_municipio)
    return pop[["ano", "cod_ibge", "municipio", "uf", "populacao"]].drop_duplicates()


def extrair_receita_corrente(receitas: pd.DataFrame) -> pd.DataFrame:
    filtro = (
        receitas["coluna"].eq("Receitas Brutas Realizadas")
        & receitas["conta"].astype(str).str.startswith("1.0.0.0.00.0.0")
    )
    return (
        receitas.loc[filtro]
        .groupby(["ano", "cod_ibge"], as_index=False)["valor"]
        .sum()
        .rename(columns={"valor": "receita_corrente"})
    )


def extrair_despesas(despesas: pd.DataFrame) -> pd.DataFrame:
    empenhadas = despesas["coluna"].eq("Despesas Empenhadas")
    pessoal = despesas["conta"].astype(str).str.startswith("3.1.00.00.00")
    investimento = despesas["conta"].astype(str).str.startswith("4.4.00.00.00")

    desp_pessoal = (
        despesas.loc[empenhadas & pessoal]
        .groupby(["ano", "cod_ibge"], as_index=False)["valor"]
        .sum()
        .rename(columns={"valor": "despesa_pessoal"})
    )
    desp_investimento = (
        despesas.loc[empenhadas & investimento]
        .groupby(["ano", "cod_ibge"], as_index=False)["valor"]
        .sum()
        .rename(columns={"valor": "investimento"})
    )
    return desp_pessoal.merge(desp_investimento, on=["ano", "cod_ibge"], how="outer")


def montar_base_bruta(
    receitas: pd.DataFrame, despesas: pd.DataFrame, pib: pd.DataFrame
) -> pd.DataFrame:
    populacao = preparar_populacao(receitas, despesas)
    receita_corrente = extrair_receita_corrente(receitas)
    despesa = extrair_despesas(despesas)

    base = (
        populacao.merge(receita_corrente, on=["ano", "cod_ibge"], how="left")
        .merge(despesa, on=["ano", "cod_ibge"], how="left")
        .merge(pib, on=["ano", "cod_ibge"], how="left")
    )
    base["municipio"] = base["municipio"].fillna(base["municipio_pib"])
    return base


def marcar_pequenos_municipios(base: pd.DataFrame) -> pd.DataFrame:
    base = base.copy()
    referencia = (
        base.loc[
            base["ano"].eq(CONFIG.ano_referencia_pequenos),
            ["cod_ibge", "populacao"],
        ]
        .dropna()
        .drop_duplicates("cod_ibge")
        .sort_values("populacao")
    )
    quantidade_pequenos = int(np.ceil(len(referencia) / 4))
    pequenos = set(referencia.head(quantidade_pequenos)["cod_ibge"].astype(int))
    base["dummy_pequeno_municipio"] = (
        base["cod_ibge"].astype(int).isin(pequenos).astype(int)
    )
    return base


def calcular_variaveis_modelo(base: pd.DataFrame) -> pd.DataFrame:
    base = base.sort_values(["cod_ibge", "ano"]).copy()
    base["pib_per_capita"] = base["pib"] / base["populacao"]
    base["log_pib_per_capita"] = np.log(base["pib_per_capita"])
    base["investimento_pib"] = base["investimento"] / base["pib"]
    base["gasto_pessoal_pib"] = base["despesa_pessoal"] / base["pib"]
    base["receita_corrente_pib"] = base["receita_corrente"] / base["pib"]
    base["crescimento_populacao"] = base.groupby("cod_ibge")["populacao"].pct_change()
    base["interacao_pequeno_pessoal"] = (
        base["dummy_pequeno_municipio"] * base["gasto_pessoal_pib"]
    )

    if "distancia_capital_km" not in base.columns:
        base["distancia_capital_km"] = np.nan
    if "area_km2" not in base.columns:
        base["area_km2"] = np.nan
    base["area_por_habitante"] = base["area_km2"] / base["populacao"]
    return base


def integrar_area(base: pd.DataFrame, area: pd.DataFrame | None) -> pd.DataFrame:
    if area is None or area.empty:
        base = base.copy()
        base["area_km2"] = np.nan
        return base

    area = area[["cod_ibge", "area_km2"]].drop_duplicates("cod_ibge")
    return base.merge(area, on="cod_ibge", how="left")


def integrar_distancia(base: pd.DataFrame, distancia: pd.DataFrame | None) -> pd.DataFrame:
    if distancia is None or distancia.empty:
        base = base.copy()
        base["distancia_capital_km"] = np.nan
        return base

    distancia = distancia[["cod_ibge", "distancia_capital_km"]].drop_duplicates("cod_ibge")
    return base.merge(distancia, on="cod_ibge", how="left")


def integrar_populacao_externa(
    base: pd.DataFrame, populacao_externa: pd.DataFrame
) -> pd.DataFrame:
    pop = (
        populacao_externa[["ano", "cod_ibge", "populacao"]]
        .copy()
        .assign(
            cod_ibge=lambda d: d["cod_ibge"].astype("Int64"),
            ano=lambda d: d["ano"].astype("Int64"),
        )
    )
    check = base.merge(pop, on=["ano", "cod_ibge"], suffixes=("_siconfi", "_externa"))
    mascara = (
        check["populacao_siconfi"].notna()
        & check["populacao_externa"].notna()
        & (
            (check["populacao_siconfi"] - check["populacao_externa"]).abs()
            / check["populacao_siconfi"].abs()
            > 0.05
        )
    )
    n_div = int(mascara.sum())
    if n_div > 0:
        print(
            "  [AVISO] "
            f"{n_div} obs com divergencia >5% entre SICONFI e populacao externa"
        )

    base = base.drop(columns=["populacao"])
    return base.merge(pop, on=["ano", "cod_ibge"], how="left")


def preparar_base_final(
    receitas: pd.DataFrame,
    despesas: pd.DataFrame,
    pib: pd.DataFrame,
    area: pd.DataFrame | None = None,
    distancia: pd.DataFrame | None = None,
    populacao: pd.DataFrame | None = None,
) -> pd.DataFrame:
    base = montar_base_bruta(receitas, despesas, pib)
    base = integrar_area(base, area)
    base = integrar_distancia(base, distancia)
    if populacao is not None:
        base = integrar_populacao_externa(base, populacao)
    base = marcar_pequenos_municipios(base)
    base = calcular_variaveis_modelo(base)
    base = base[
        base["ano"].between(CONFIG.ano_inicial, CONFIG.ano_final)
    ].copy()
    return base[COLUNAS_BASE_FINAL].sort_values(["ano", "municipio"])
