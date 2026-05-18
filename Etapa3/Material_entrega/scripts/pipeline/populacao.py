from __future__ import annotations

import gzip
import http.cookiejar
import json
import re
import tempfile
import time
from html.parser import HTMLParser
from io import StringIO
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, build_opener, HTTPCookieProcessor, urlopen

import pandas as pd

from pipeline.config import CONFIG, PATHS

# ── Constantes ───────────────────────────────────────────────────────────────

TABNET_URL = "https://tabnet.datasus.gov.br/cgi/tabcgi.exe?ibge/cnv/popsvs2024br.def"
TABNET_ENCODING = "iso-8859-1"

IBGE_FTP_BASE = "https://ftp.ibge.gov.br/Estimativas_de_Populacao/"
SIDRA_URL = "https://apisidra.ibge.gov.br/values/t/9514/n6/in%20n3%20{cod}/v/93/p/{ano}?formato=json"

FONTE_DATASUS = "DATASUS/SVS - popsvs2024br - Estimativas populacionais por municipio"
FONTE_IBGE_FTP = "IBGE FTP - estimativa_dou - Estimativas populacionais por municipio"
FONTE_SIDRA = "IBGE SIDRA t/9514 - Censo 2022 - Populacao residente"
FONTE_SICONFI = "SICONFI/Tesouro Nacional - coluna populacao (fonte IBGE)"

ANOS_DEFAULT = list(range(2018, 2026))

# Arquivos DOU disponíveis no IBGE FTP (confirmado)
# 2022 e 2023 nao estao neste diretorio
_DOU_ARQUIVOS: dict[int, str] = {
    2018: "Estimativas_2018/estimativa_dou_2018_20181019.xls",
    2019: "Estimativas_2019/estimativa_dou_2019.xls",
    2020: "Estimativas_2020/estimativa_dou_2020.xls",
    2021: "Estimativas_2021/estimativa_dou_2021.xls",
    2024: "Estimativas_2024/estimativa_dou_2024.xls",
    2025: "Estimativas_2025/estimativa_dou_2025.xls",
}

_CODIGOS_UF = {
    "RO": "11", "AC": "12", "AM": "13", "RR": "14", "PA": "15",
    "AP": "16", "TO": "17", "MA": "21", "PI": "22", "CE": "23",
    "RN": "24", "PB": "25", "PE": "26", "AL": "27", "SE": "28",
    "BA": "29", "MG": "31", "ES": "32", "RJ": "33", "SP": "35",
    "PR": "41", "SC": "42", "RS": "43", "MS": "50", "MT": "51",
    "GO": "52", "DF": "53",
}

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Content-Type": "application/x-www-form-urlencoded",
    "Referer": TABNET_URL,
}


def prefixo_ibge(uf: str) -> str:
    try:
        return _CODIGOS_UF[uf.upper()]
    except KeyError as exc:
        raise ValueError(f"UF nao reconhecida: {uf}") from exc


# ── HTTP / HTTP com cookies ──────────────────────────────────────────────────

_jar = http.cookiejar.CookieJar()
_opener = build_opener(HTTPCookieProcessor(_jar))


def _fetch_tabnet(url: str, data: bytes | None = None, timeout: int = 90) -> bytes:
    req = Request(url, data=data, headers=_HEADERS)
    with _opener.open(req, timeout=timeout) as resp:
        raw = resp.read()
    try:
        return gzip.decompress(raw)
    except OSError:
        return raw


def _fetch_ibge(url: str, timeout: int = 60) -> bytes:
    with urlopen(url, timeout=timeout) as resp:
        return resp.read()


# ── Fontes de dados ──────────────────────────────────────────────────────────

# -- DATASUS TabNet -----------------------------------------------------------

class _FormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.selects: dict[str, list[str]] = {}
        self.hidden: dict[str, str] = {}
        self.checkboxes: list[tuple[str, str]] = []
        self._sel: str | None = None

    def handle_starttag(self, tag: str, attrs: list) -> None:
        a = dict(attrs)
        if tag == "input":
            tipo = a.get("type", "text").lower()
            nm, val = a.get("name", ""), a.get("value", "")
            if tipo == "hidden" and nm:
                self.hidden[nm] = val
            elif tipo in ("checkbox", "radio") and nm:
                self.checkboxes.append((nm, val))
        elif tag == "select":
            nm = a.get("name")
            if nm:
                self._sel = nm
                self.selects[nm] = []
        elif tag == "option" and self._sel:
            val = a.get("value", "")
            if val:
                self.selects[self._sel].append(val)

    def handle_endtag(self, tag: str) -> None:
        if tag == "select":
            self._sel = None


def _match_opcao(opcoes: list[str], *termos: str) -> str | None:
    termos_l = [t.lower() for t in termos]
    for op in opcoes:
        if all(t in op.lower() for t in termos_l):
            return op
    return None


def _montar_payload_tabnet(form: _FormParser, anos: list[int]) -> bytes:
    dados: dict[str, str] = {}
    dados.update(form.hidden)

    linha_opts = form.selects.get("Linha", [])
    dados["Linha"] = _match_opcao(linha_opts, "munic") or (linha_opts[0] if linha_opts else "Município")

    col_opts = form.selects.get("Coluna", [])
    dados["Coluna"] = _match_opcao(col_opts, "ano") or (col_opts[0] if col_opts else "Ano")

    incr_opts = form.selects.get("Incremento", [])
    dados["Incremento"] = incr_opts[0] if incr_opts else "Pop_resid"

    anos_str = {str(a) for a in anos}
    for nm, val in form.checkboxes:
        nm_digits = re.sub(r"[^0-9]", "", nm)
        val_digits = re.sub(r"[^0-9]", "", val)
        if nm_digits in anos_str or val_digits in anos_str:
            dados[nm] = val if val else "on"

    dados["mostre"] = "Mostra"
    return urlencode(dados, encoding="iso-8859-1").encode("iso-8859-1")


def _parse_tabela_tabnet(html: bytes, uf: str) -> pd.DataFrame:
    texto = html.decode(TABNET_ENCODING, errors="replace")
    try:
        tabelas = pd.read_html(StringIO(texto), thousands=".", decimal=",")
    except ValueError as exc:
        raise ValueError(f"Nenhuma tabela HTML na resposta do TabNet: {exc}") from exc

    tabela = max(tabelas, key=lambda t: t.shape[0] * t.shape[1])
    tabela.columns = [str(c).strip() for c in tabela.columns]
    col_mun = tabela.columns[0]
    cols_ano = [c for c in tabela.columns[1:] if re.fullmatch(r"20\d{2}", c)]

    if not cols_ano:
        raise ValueError(f"Colunas de ano nao encontradas. Colunas: {list(tabela.columns[:10])}")

    tabela = tabela[
        tabela[col_mun].notna()
        & ~tabela[col_mun].astype(str).str.strip().isin(["Total", ""])
    ].copy()

    tabela["cod_ibge"] = (
        tabela[col_mun].astype(str)
        .str.extract(r"(\d{6,7})", expand=False)
        .pipe(pd.to_numeric, errors="coerce")
        .astype("Int64")
    )
    tabela["municipio"] = (
        tabela[col_mun].astype(str)
        .str.replace(r"^\d+\s*", "", regex=True)
        .str.strip()
    )

    pref = prefixo_ibge(uf)
    tabela = tabela[
        tabela["cod_ibge"].notna()
        & tabela["cod_ibge"].astype(str).str.startswith(pref)
    ].copy()

    if tabela.empty:
        raise ValueError(f"Nenhum municipio de {uf} na resposta (prefixo '{pref}').")

    longo = tabela.melt(
        id_vars=["cod_ibge", "municipio"],
        value_vars=cols_ano,
        var_name="ano",
        value_name="populacao",
    )
    longo["ano"] = pd.to_numeric(longo["ano"], errors="coerce").astype("Int64")
    longo["populacao"] = (
        longo["populacao"].astype(str)
        .str.replace(r"[^\d]", "", regex=True)
        .replace("", pd.NA)
        .pipe(pd.to_numeric, errors="coerce")
        .astype("Int64")
    )
    longo["uf"] = uf.upper()
    longo["fonte"] = FONTE_DATASUS
    return (
        longo[["cod_ibge", "municipio", "uf", "ano", "populacao", "fonte"]]
        .dropna(subset=["cod_ibge", "ano", "populacao"])
        .sort_values(["ano", "cod_ibge"])
        .reset_index(drop=True)
    )


def baixar_populacao_datasus(
    anos: list[int] = ANOS_DEFAULT,
    uf: str = CONFIG.uf,
    timeout: int = 90,
) -> pd.DataFrame:
    """Baixa estimativas populacionais do DATASUS TabNet (requer acesso a rede DATASUS)."""
    html_form = _fetch_tabnet(TABNET_URL, timeout=timeout)
    form = _FormParser()
    form.feed(html_form.decode(TABNET_ENCODING, errors="replace"))

    if not form.selects:
        raise RuntimeError(
            "Formulario TabNet nao parseado corretamente. "
            "Verifique conectividade com tabnet.datasus.gov.br"
        )

    payload = _montar_payload_tabnet(form, anos)
    time.sleep(1)
    html_result = _fetch_tabnet(TABNET_URL, data=payload, timeout=timeout)
    return _parse_tabela_tabnet(html_result, uf)


# -- IBGE FTP (DOU) ----------------------------------------------------------

def _aba_municipios(xls: pd.ExcelFile) -> str:
    """Retorna o nome da aba de municipios, independente de encoding."""
    for aba in xls.sheet_names:
        if "MUNIC" in aba.upper():
            return aba
    raise ValueError(f"Aba de municipios nao encontrada. Abas: {xls.sheet_names}")


def _baixar_dou_ibge(ano: int, uf: str) -> pd.DataFrame:
    caminho_relativo = _DOU_ARQUIVOS[ano]
    url = IBGE_FTP_BASE + caminho_relativo
    raw = _fetch_ibge(url)

    tmp = tempfile.NamedTemporaryFile(suffix=".xls", delete=False)
    tmp.write(raw)
    tmp.close()  # fecha antes de abrir com pandas (necessario no Windows)
    tmp_path = tmp.name

    try:
        xls = pd.ExcelFile(tmp_path)
        aba = _aba_municipios(xls)
        df = pd.read_excel(xls, sheet_name=aba, skiprows=1)
        xls.close()
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    # Normaliza: algumas versoes tem 5 colunas, outras 6 (extra vazia)
    df = df.iloc[:, :5].copy()
    df.columns = ["uf", "cod_uf", "cod_munic", "municipio", "populacao"]

    df = df[df["uf"].notna() & (df["uf"].astype(str).str.len() == 2)].copy()
    df["cod_uf"] = pd.to_numeric(df["cod_uf"], errors="coerce").astype("Int64")
    df["cod_munic"] = pd.to_numeric(df["cod_munic"], errors="coerce").astype("Int64")
    df = df[df["cod_uf"].notna() & df["cod_munic"].notna()].copy()

    df["cod_ibge"] = (
        df["cod_uf"].astype(int).astype(str).str.zfill(2)
        + df["cod_munic"].astype(int).astype(str).str.zfill(5)
    ).astype(int)
    df["populacao"] = pd.to_numeric(df["populacao"], errors="coerce").astype("Int64")
    df["ano"] = ano
    df["fonte"] = FONTE_IBGE_FTP

    df = df[df["uf"].astype(str).str.strip().eq(uf.upper())].copy()
    if df.empty:
        raise ValueError(f"Nenhum municipio de {uf} no arquivo DOU {ano}.")

    return df[["cod_ibge", "municipio", "uf", "ano", "populacao", "fonte"]]


# -- IBGE SIDRA (Censo 2022) -------------------------------------------------

def _baixar_sidra_censo(ano: int, uf: str) -> pd.DataFrame:
    """Baixa populacao via SIDRA t/9514 (Censo 2022). Disponivel apenas para 2022."""
    cod = prefixo_ibge(uf)
    url = SIDRA_URL.format(cod=cod, ano=ano)
    with urlopen(url, timeout=30) as r:
        dados = json.loads(r.read().decode("utf-8-sig"))

    if len(dados) <= 1:
        raise ValueError(f"SIDRA nao retornou dados para {uf} em {ano}.")

    registros = dados[1:]
    rows = []
    for d in registros:
        cod_ibge_str = d.get("D1C", "")
        m = re.search(r"\d{7}", cod_ibge_str)
        if not m:
            continue
        rows.append({
            "cod_ibge": int(m.group()),
            "municipio": re.sub(r"\s*-\s*[A-Z]{2}$", "", d.get("D1N", "")),
            "uf": uf.upper(),
            "ano": ano,
            "populacao": pd.to_numeric(d.get("V"), errors="coerce"),
            "fonte": FONTE_SIDRA,
        })

    df = pd.DataFrame(rows)
    df["cod_ibge"] = df["cod_ibge"].astype("Int64")
    df["populacao"] = df["populacao"].astype("Int64")
    return df[["cod_ibge", "municipio", "uf", "ano", "populacao", "fonte"]]


# -- SICONFI (extração local) ------------------------------------------------

def _extrair_siconfi(ano: int, uf: str) -> pd.DataFrame:
    """
    Extrai populacao dos arquivos SICONFI (Receitas) ja presentes no projeto.
    Coluna 'populacao' nesses arquivos vem do IBGE (repassada pelo municipio ao SICONFI).
    """
    pasta = PATHS.receitas
    candidatos = sorted(pasta.glob(f"*{ano}*.xls*"))
    if not candidatos:
        raise FileNotFoundError(
            f"Arquivo SICONFI de receitas para {ano} nao encontrado em {pasta}."
        )

    df = pd.read_excel(candidatos[0], skiprows=3)
    # Normaliza nomes
    df.columns = [
        c.strip().lower().replace(" ", "_").replace(".", "").replace("/", "")
        for c in df.columns
    ]

    col_ibge = next((c for c in df.columns if "ibge" in c or "cod" in c and "ibge" in c), None)
    col_pop = next((c for c in df.columns if "popul" in c), None)
    col_inst = next((c for c in df.columns if "instit" in c), None)
    col_uf = next((c for c in df.columns if c == "uf" or c == "sg_uf"), None)

    if not col_pop:
        raise ValueError(f"Coluna de populacao nao encontrada no SICONFI {ano}.")

    filtro = df[col_uf].astype(str).str.strip().eq(uf.upper()) if col_uf else slice(None)
    df = df[filtro].copy()

    df["cod_ibge"] = pd.to_numeric(df[col_ibge], errors="coerce").astype("Int64") if col_ibge else pd.NA
    df["populacao"] = pd.to_numeric(df[col_pop], errors="coerce").astype("Int64")

    if col_inst:
        df["municipio"] = (
            df[col_inst].astype(str)
            .str.replace(r"^Prefeitura Municipal de\s+", "", regex=True)
            .str.replace(r"\s*-\s*[A-Z]{2}$", "", regex=True)
            .str.strip()
        )
    else:
        df["municipio"] = ""

    df["ano"] = ano
    df["uf"] = uf.upper()
    df["fonte"] = FONTE_SICONFI

    resultado = (
        df[["cod_ibge", "municipio", "uf", "ano", "populacao", "fonte"]]
        .dropna(subset=["cod_ibge", "populacao"])
        .drop_duplicates("cod_ibge")
        .sort_values("cod_ibge")
        .reset_index(drop=True)
    )

    if resultado.empty:
        raise ValueError(f"Nenhum municipio de {uf} extraido do SICONFI {ano}.")
    return resultado


# ── Orquestrador com fallback ────────────────────────────────────────────────

def baixar_populacao_ibge(
    anos: list[int] = ANOS_DEFAULT,
    uf: str = CONFIG.uf,
) -> pd.DataFrame:
    """
    Fallback ao DATASUS: baixa dados do IBGE FTP + SIDRA + SICONFI.

    Cobertura confirmada:
      2018, 2019, 2020, 2021 → IBGE FTP (estimativa_dou_*.xls)
      2022                   → IBGE SIDRA tabela 9514 (Censo 2022)
      2023                   → SICONFI (Receitas Orcamentarias - 2023.xlsx)
      2024, 2025             → IBGE FTP (estimativa_dou_*.xls)
    """
    frames = []

    for ano in anos:
        try:
            if ano in _DOU_ARQUIVOS:
                print(f"    {ano}: IBGE FTP...")
                df = _baixar_dou_ibge(ano, uf)
            elif ano == 2022:
                print(f"    {ano}: IBGE SIDRA (Censo)...")
                df = _baixar_sidra_censo(ano, uf)
            else:
                print(f"    {ano}: SICONFI (arquivo local)...")
                df = _extrair_siconfi(ano, uf)
            frames.append(df)
        except Exception as exc:
            print(f"    [AVISO] Ano {ano} nao obtido: {exc}")

    if not frames:
        raise RuntimeError("Nenhum dado foi obtido de nenhuma fonte.")

    return (
        pd.concat(frames, ignore_index=True)
        .sort_values(["ano", "cod_ibge"])
        .reset_index(drop=True)
    )


# ── Cache e persistência ─────────────────────────────────────────────────────

def _caminho_combinado(uf: str = CONFIG.uf) -> Path:
    return PATHS.base_dados / "populacao" / f"populacao_{uf.lower()}_2018_2025.csv"


def _caminho_ano(ano: int, uf: str = CONFIG.uf) -> Path:
    return PATHS.base_dados / "populacao" / f"populacao_{uf.lower()}_{ano}.csv"


def salvar_populacao(df: pd.DataFrame, uf: str = CONFIG.uf) -> list[Path]:
    """Salva um CSV por ano + CSV combinado em base_dados/populacao/."""
    pasta = PATHS.base_dados / "populacao"
    pasta.mkdir(parents=True, exist_ok=True)

    salvos: list[Path] = []
    for ano_val, grupo in df.groupby("ano"):
        cam = _caminho_ano(int(ano_val), uf)
        grupo.to_csv(
            cam,
            index=False,
            sep=CONFIG.csv_sep,
            decimal=CONFIG.csv_decimal,
            encoding=CONFIG.csv_encoding,
        )
        salvos.append(cam)
        print(f"    salvo: {cam.name}  ({len(grupo)} municipios)")

    cam_comb = _caminho_combinado(uf)
    df.to_csv(
        cam_comb,
        index=False,
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
        encoding=CONFIG.csv_encoding,
    )
    salvos.append(cam_comb)
    return salvos


def ler_populacao(uf: str = CONFIG.uf) -> pd.DataFrame:
    return pd.read_csv(
        _caminho_combinado(uf),
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
        encoding="utf-8-sig",
    )


def obter_populacao(
    anos: list[int] = ANOS_DEFAULT,
    uf: str = CONFIG.uf,
    force_download: bool = False,
    usar_datasus: bool = True,
) -> pd.DataFrame:
    """
    Padrao cache. Parametros:
      force_download: forca re-download mesmo com cache.
      usar_datasus: tenta DATASUS antes do IBGE FTP (padrao True).
    """
    if _caminho_combinado(uf).exists() and not force_download:
        return ler_populacao(uf)

    if usar_datasus:
        try:
            print("  Tentando DATASUS TabNet...")
            df = baixar_populacao_datasus(anos, uf)
            salvar_populacao(df, uf)
            return df
        except Exception as exc:
            print(f"  DATASUS inacessivel ({exc}). Usando IBGE FTP/SIDRA/SICONFI...")

    df = baixar_populacao_ibge(anos, uf)
    salvar_populacao(df, uf)
    return df
