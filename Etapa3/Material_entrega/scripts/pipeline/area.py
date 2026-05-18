from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import quote
from urllib.request import urlopen

import pandas as pd

from pipeline.config import CONFIG, PATHS


FONTE_AREA = "SIDRA/IBGE tabela 4714, variavel 6318"
SIDRA_AREA_URL = (
    "https://apisidra.ibge.gov.br/values/"
    "t/4714/n6/in%20n3%20{uf_codigo}/v/6318/p/last?formato=json"
)
UF_CODIGOS = {
    "RO": 11,
    "AC": 12,
    "AM": 13,
    "RR": 14,
    "PA": 15,
    "AP": 16,
    "TO": 17,
    "MA": 21,
    "PI": 22,
    "CE": 23,
    "RN": 24,
    "PB": 25,
    "PE": 26,
    "AL": 27,
    "SE": 28,
    "BA": 29,
    "MG": 31,
    "ES": 32,
    "RJ": 33,
    "SP": 35,
    "PR": 41,
    "SC": 42,
    "RS": 43,
    "MS": 50,
    "MT": 51,
    "GO": 52,
    "DF": 53,
}


def codigo_uf(uf: str) -> int:
    try:
        return UF_CODIGOS[uf.upper()]
    except KeyError as exc:
        raise ValueError(f"UF nao cadastrada para consulta SIDRA: {uf}") from exc


def caminho_area_complementar(uf: str = CONFIG.uf) -> Path:
    return PATHS.outputs / f"base_area_municipios_{uf.lower()}.csv"


def url_area_municipios(uf: str) -> str:
    return SIDRA_AREA_URL.format(uf_codigo=quote(str(codigo_uf(uf))))


def consultar_area_municipios(uf: str, timeout: int = 60) -> list[dict]:
    with urlopen(url_area_municipios(uf), timeout=timeout) as response:
        dados = json.loads(response.read().decode("utf-8-sig"))

    if len(dados) <= 1:
        raise ValueError("A consulta ao SIDRA nao retornou municipios.")
    return dados[1:]


def parse_area_sidra(registros: list[dict], uf: str = CONFIG.uf) -> pd.DataFrame:
    df = pd.DataFrame(registros)
    area = pd.DataFrame(
        {
            "cod_ibge": pd.to_numeric(df["D1C"], errors="coerce").astype("Int64"),
            "municipio_area": df["D1N"]
            .astype(str)
            .str.replace(r"\s+-\s+[A-Z]{2}$", "", regex=True),
            "uf": uf.upper(),
            "area_km2": pd.to_numeric(df["V"], errors="coerce"),
            "ano_area": pd.to_numeric(df["D3C"], errors="coerce").astype("Int64"),
            "fonte_area": FONTE_AREA,
        }
    )
    return area.dropna(subset=["cod_ibge", "area_km2"]).sort_values("municipio_area")


def baixar_area_sidra(uf: str = CONFIG.uf) -> pd.DataFrame:
    registros = consultar_area_municipios(uf)
    return parse_area_sidra(registros, uf)


def salvar_area_complementar(area: pd.DataFrame, uf: str = CONFIG.uf) -> Path:
    PATHS.outputs.mkdir(parents=True, exist_ok=True)
    caminho = caminho_area_complementar(uf)
    area.to_csv(
        caminho,
        index=False,
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
        encoding=CONFIG.csv_encoding,
    )
    return caminho


def ler_area_complementar(uf: str = CONFIG.uf) -> pd.DataFrame:
    return pd.read_csv(
        caminho_area_complementar(uf),
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
    )


def obter_area_municipios(force_download: bool = False) -> pd.DataFrame:
    if caminho_area_complementar(CONFIG.uf).exists() and not force_download:
        return ler_area_complementar()

    area = baixar_area_sidra(CONFIG.uf)
    salvar_area_complementar(area, CONFIG.uf)
    return area
