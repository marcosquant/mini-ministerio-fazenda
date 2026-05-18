from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ProjectPaths:
    root: Path = field(default_factory=lambda: Path(__file__).resolve().parents[3])

    @property
    def base_dados(self) -> Path:
        return self.root / "base_dados"

    @property
    def outputs(self) -> Path:
        return self.root / "Material_entrega" / "outputs"

    @property
    def receitas(self) -> Path:
        return self.base_dados / "Receitas"

    @property
    def despesas(self) -> Path:
        return self.base_dados / "Despesas"

    @property
    def pib(self) -> Path:
        return self.base_dados / "PIB"

    @property
    def populacao(self) -> Path:
        return self.base_dados / "populacao"


@dataclass(frozen=True)
class PipelineConfig:
    uf: str = "MS"
    ano_inicial: int = 2019
    ano_final: int = 2023
    ano_referencia_pequenos: int = 2019
    output_base: str = "base_etapa3_ms.csv"
    csv_sep: str = ";"
    csv_decimal: str = ","
    csv_encoding: str = "utf-8-sig"


PATHS = ProjectPaths()
CONFIG = PipelineConfig()


COLUNAS_BASE_FINAL = [
    "ano",
    "cod_ibge",
    "municipio",
    "uf",
    "populacao",
    "pib",
    "pib_per_capita",
    "pib_per_capita_ibge",
    "receita_corrente",
    "despesa_pessoal",
    "investimento",
    "crescimento_pib",
    "crescimento_populacao",
    "log_pib_per_capita",
    "receita_corrente_pib",
    "gasto_pessoal_pib",
    "investimento_pib",
    "dummy_pequeno_municipio",
    "interacao_pequeno_pessoal",
    "area_km2",
    "area_por_habitante",
    "distancia_capital_km",
]
