from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd


def extrair_ano(caminho: Path) -> int:
    match = re.search(r"(20\d{2})", caminho.name)
    if not match:
        raise ValueError(f"Nao foi possivel identificar o ano no arquivo: {caminho}")
    return int(match.group(1))


def listar_excel(pasta: Path, extensoes: tuple[str, ...] = (".xlsx", ".xls")) -> list[Path]:
    return sorted(
        p
        for p in pasta.iterdir()
        if p.is_file() and p.suffix.lower() in extensoes and not p.name.startswith("~$")
    )


def normalizar_texto(texto: object) -> str:
    sem_acento = unicodedata.normalize("NFKD", str(texto))
    sem_acento = "".join(char for char in sem_acento if not unicodedata.combining(char))
    return " ".join(sem_acento.lower().split())


def localizar_coluna(df: pd.DataFrame, partes: list[str]) -> str:
    partes_normalizadas = [normalizar_texto(parte) for parte in partes]
    for coluna in df.columns:
        texto = normalizar_texto(coluna)
        if all(parte in texto for parte in partes_normalizadas):
            return coluna
    raise KeyError(f"Coluna nao encontrada com partes: {partes}")


def renomear_colunas_normalizadas(
    df: pd.DataFrame, mapa_normalizado: dict[str, str]
) -> pd.DataFrame:
    renomear = {}
    for coluna in df.columns:
        chave = normalizar_texto(coluna)
        if chave in mapa_normalizado:
            renomear[coluna] = mapa_normalizado[chave]
    return df.rename(columns=renomear)


def imprimir_resumo_base(base: pd.DataFrame, caminho_saida: Path) -> None:
    print("Base preparada com sucesso:")
    print(f"- linhas: {len(base)}")
    print(f"- municipios: {base['cod_ibge'].nunique()}")
    print(f"- anos: {int(base['ano'].min())}-{int(base['ano'].max())}")
    print(f"- arquivo: {caminho_saida}")
    print("\nValores ausentes por coluna:")
    print(base.isna().sum().to_string())
