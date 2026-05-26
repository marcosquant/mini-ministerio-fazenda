from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from pipeline.config import CONFIG, PATHS

ARQUIVO_POPULACAO_TABNET = (
    PATHS.populacao / "ibge_cnv_popsvs2024br194401201_17_157_120.csv"
)
FONTE_POPULACAO = "DATASUS/TABNET - popsvs2024br - arquivo CSV indicado pelo professor"
ANOS_DEFAULT = list(range(2018, 2024))


def _ler_csv_tabnet(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo de populacao nao encontrado: {caminho}")

    ultimo_erro: UnicodeDecodeError | None = None
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(caminho, sep=";", skiprows=4, encoding=encoding)
        except UnicodeDecodeError as exc:
            ultimo_erro = exc

    raise ultimo_erro or UnicodeDecodeError("utf-8", b"", 0, 1, "encoding invalido")


def _normalizar_nome(nome: object) -> str:
    texto = str(nome).strip()
    texto = re.sub(r"^\d+\s*", "", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def _arquivos_receitas() -> list[Path]:
    return sorted(
        caminho
        for caminho in PATHS.receitas.glob("*.xls*")
        if not caminho.name.startswith("~$")
    )


def _mapear_codigos_municipios(uf: str = CONFIG.uf) -> pd.DataFrame:
    partes: list[pd.DataFrame] = []

    for caminho in _arquivos_receitas():
        df = pd.read_excel(
            caminho,
            skiprows=3,
            usecols=["Cod.IBGE", "Instituição", "UF"],
        )
        df = df[df["UF"].astype(str).str.upper().eq(uf.upper())].copy()
        df["cod_ibge"] = pd.to_numeric(df["Cod.IBGE"], errors="coerce").astype("Int64")
        df = df.dropna(subset=["cod_ibge"])
        df["codigo_6"] = df["cod_ibge"].astype(str).str[:6]
        df["municipio"] = df["Instituição"].map(_normalizar_nome)
        partes.append(df[["codigo_6", "cod_ibge", "municipio"]])

    if not partes:
        raise FileNotFoundError(f"Nenhum arquivo de Receitas encontrado em {PATHS.receitas}")

    mapa = pd.concat(partes, ignore_index=True).drop_duplicates()
    duplicados = mapa[mapa.duplicated("codigo_6", keep=False)]
    if not duplicados.empty:
        conflitos = duplicados.sort_values("codigo_6").to_string(index=False)
        raise ValueError(f"Codigos municipais duplicados no mapeamento:\n{conflitos}")

    return mapa


def _converter_populacao(serie: pd.Series) -> pd.Series:
    valores = pd.to_numeric(serie, errors="coerce")
    faltantes = valores.isna() & serie.notna()

    if faltantes.any():
        texto = (
            serie.loc[faltantes]
            .astype(str)
            .str.strip()
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        valores.loc[faltantes] = pd.to_numeric(texto, errors="coerce")

    return valores.round().astype("Int64")


def obter_populacao(
    anos: list[int] | None = None,
    uf: str = CONFIG.uf,
    caminho: Path = ARQUIVO_POPULACAO_TABNET,
    **_: object,
) -> pd.DataFrame:
    """
    Le a populacao municipal do CSV exportado do TabNet/DATASUS.

    O arquivo indicado pelo professor possui codigos municipais com 6 digitos.
    A base final trabalha com o codigo IBGE de 7 digitos, por isso o ultimo
    digito e recuperado a partir dos codigos oficiais presentes nos arquivos de
    Receitas usados no mesmo pipeline.
    """
    anos = anos or ANOS_DEFAULT
    colunas_ano = [str(ano) for ano in anos]

    df = _ler_csv_tabnet(caminho)
    df.columns = [str(coluna).strip().strip('"') for coluna in df.columns]

    if "Município" not in df.columns:
        raise ValueError(f"Coluna 'Município' nao encontrada em {caminho.name}.")

    colunas_disponiveis = [coluna for coluna in colunas_ano if coluna in df.columns]
    if not colunas_disponiveis:
        raise ValueError(
            f"Nenhuma coluna de ano {colunas_ano} encontrada em {caminho.name}."
        )

    df = df[df["Município"].notna()].copy()
    df["codigo_6"] = df["Município"].astype(str).str.extract(r"^(\d{6})", expand=False)
    df["municipio_csv"] = df["Município"].map(_normalizar_nome)
    df = df.dropna(subset=["codigo_6"])

    mapa_codigos = _mapear_codigos_municipios(uf)
    df = df.merge(mapa_codigos, on="codigo_6", how="left")

    sem_codigo = df[df["cod_ibge"].isna()]
    if not sem_codigo.empty:
        municipios = ", ".join(sem_codigo["municipio_csv"].astype(str).head(10))
        raise ValueError(
            "Nao foi possivel converter o codigo de 6 para 7 digitos para: "
            f"{municipios}"
        )

    populacao = df.melt(
        id_vars=["cod_ibge", "municipio", "codigo_6"],
        value_vars=colunas_disponiveis,
        var_name="ano",
        value_name="populacao",
    )
    populacao["ano"] = pd.to_numeric(populacao["ano"], errors="coerce").astype("Int64")
    populacao["populacao"] = _converter_populacao(populacao["populacao"])
    populacao["uf"] = uf.upper()
    populacao["fonte"] = FONTE_POPULACAO

    populacao = (
        populacao[["cod_ibge", "municipio", "uf", "ano", "populacao", "fonte"]]
        .dropna(subset=["cod_ibge", "ano", "populacao"])
        .drop_duplicates(["cod_ibge", "ano"])
        .sort_values(["cod_ibge", "ano"])
        .reset_index(drop=True)
    )

    anos_lidos = sorted(populacao["ano"].astype(int).unique().tolist())
    print(
        "Populacao carregada do arquivo local "
        f"{caminho.name}: {len(populacao)} registros, anos {anos_lidos}."
    )

    return populacao
