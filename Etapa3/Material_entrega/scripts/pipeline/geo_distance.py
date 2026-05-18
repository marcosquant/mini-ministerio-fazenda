from __future__ import annotations

from pathlib import Path

import pandas as pd

from pipeline.config import CONFIG, PATHS


CODIGO_CAMPO_GRANDE = 5002704
NOME_CAPITAL = "Campo Grande"
FONTE_DISTANCIA = "REGIC 2018 - Rotas Brasil"
REGIC_PATH = (
    PATHS.base_dados
    / "distancias-municipios-capital"
    / "REGIC2018_Rotas_Brasil.xlsx"
)
REGIC_SHEET = "rotas_regic2018_brasil"
REGIC_COLUMNS = [
    "modal",
    "sig_uf_o",
    "cod_o",
    "nome_o",
    "sig_uf_d",
    "cod_d",
    "nome_d",
    "km",
    "minutos",
]
OUTPUT_COLUMNS = [
    "cod_ibge",
    "municipio",
    "uf",
    "distancia_capital_km",
    "minutos_capital",
    "modal_distancia",
    "metodo_distancia",
    "fonte_distancia",
]


def caminho_distancia_complementar(uf: str = CONFIG.uf) -> Path:
    return PATHS.outputs / f"base_distancia_capital_{uf.lower()}.csv"


def ler_rotas_regic(caminho: Path = REGIC_PATH) -> pd.DataFrame:
    rotas = pd.read_excel(caminho, sheet_name=REGIC_SHEET, usecols=REGIC_COLUMNS)
    for coluna in ["cod_o", "cod_d"]:
        rotas[coluna] = pd.to_numeric(rotas[coluna], errors="coerce").astype("Int64")
    for coluna in ["km", "minutos"]:
        rotas[coluna] = pd.to_numeric(rotas[coluna], errors="coerce")
    return rotas


def rotas_municipios_para_capital(rotas: pd.DataFrame, uf: str = CONFIG.uf) -> pd.DataFrame:
    uf = uf.upper()
    destino_capital = rotas["cod_d"].eq(CODIGO_CAMPO_GRANDE) & rotas["sig_uf_o"].eq(uf)
    origem_capital = rotas["cod_o"].eq(CODIGO_CAMPO_GRANDE) & rotas["sig_uf_d"].eq(uf)

    ida = rotas.loc[destino_capital, ["cod_o", "nome_o", "km", "minutos", "modal"]].copy()
    ida.columns = ["cod_ibge", "municipio", "distancia_capital_km", "minutos_capital", "modal_distancia"]

    volta = rotas.loc[origem_capital, ["cod_d", "nome_d", "km", "minutos", "modal"]].copy()
    volta.columns = ["cod_ibge", "municipio", "distancia_capital_km", "minutos_capital", "modal_distancia"]

    distancias = pd.concat([ida, volta], ignore_index=True)
    distancias = (
        distancias.dropna(subset=["cod_ibge", "distancia_capital_km"])
        .sort_values(["cod_ibge", "distancia_capital_km"])
        .drop_duplicates("cod_ibge")
    )
    return distancias


def adicionar_capital(distancias: pd.DataFrame) -> pd.DataFrame:
    capital = pd.DataFrame(
        [
            {
                "cod_ibge": CODIGO_CAMPO_GRANDE,
                "municipio": NOME_CAPITAL,
                "distancia_capital_km": 0.0,
                "minutos_capital": 0,
                "modal_distancia": "Capital",
            }
        ]
    )
    return pd.concat([distancias, capital], ignore_index=True)


def montar_distancias_capital(uf: str = CONFIG.uf) -> pd.DataFrame:
    distancias = rotas_municipios_para_capital(ler_rotas_regic(), uf)
    distancias = adicionar_capital(distancias)
    distancias = (
        distancias.sort_values(["cod_ibge", "distancia_capital_km"])
        .drop_duplicates("cod_ibge")
        .sort_values("cod_ibge")
    )
    distancias["uf"] = uf.upper()
    distancias["metodo_distancia"] = "Distancia de deslocamento ate a capital pela base REGIC"
    distancias["fonte_distancia"] = FONTE_DISTANCIA
    return distancias[OUTPUT_COLUMNS]


def salvar_distancia_complementar(distancias: pd.DataFrame, uf: str = CONFIG.uf) -> Path:
    PATHS.outputs.mkdir(parents=True, exist_ok=True)
    caminho = caminho_distancia_complementar(uf)
    distancias.to_csv(
        caminho,
        index=False,
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
        encoding=CONFIG.csv_encoding,
    )
    return caminho


def obter_distancia_capital() -> pd.DataFrame:
    distancias = montar_distancias_capital(CONFIG.uf)
    salvar_distancia_complementar(distancias, CONFIG.uf)
    return distancias
