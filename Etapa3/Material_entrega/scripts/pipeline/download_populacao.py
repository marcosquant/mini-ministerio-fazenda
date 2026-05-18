"""
Download de estimativas populacionais municipais (2018-2025).

Fontes (em ordem de prioridade):
  1. DATASUS TabNet (popsvs2024br.def) — requer acesso a tabnet.datasus.gov.br
  2. IBGE FTP (estimativa_dou_*.xls)  — 2018-2021, 2024, 2025
  3. IBGE SIDRA t/9514                — 2022 (Censo 2022)
  4. SICONFI (arquivo local)          — 2023

Uso:
    cd Etapa3/Material_entrega/scripts
    python -m pipeline.download_populacao              # baixa 2018-2025 para MS
    python -m pipeline.download_populacao --force      # forca re-download
    python -m pipeline.download_populacao --sem-datasus  # pula tentativa DATASUS

Saida em Etapa3/base_dados/populacao/:
    populacao_ms_<ano>.csv     (um por ano)
    populacao_ms_2018_2025.csv (combinado)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1]))

from pipeline.config import CONFIG
from pipeline.populacao import ANOS_DEFAULT, obter_populacao


def main(force: bool = False, sem_datasus: bool = False) -> None:
    print("=" * 60)
    print("Download populacao municipal")
    print(f"  UF    : {CONFIG.uf}")
    print(f"  Anos  : {ANOS_DEFAULT[0]} a {ANOS_DEFAULT[-1]}")
    print("=" * 60)

    try:
        df = obter_populacao(
            anos=ANOS_DEFAULT,
            uf=CONFIG.uf,
            force_download=force,
            usar_datasus=not sem_datasus,
        )
    except Exception as exc:
        print(f"\nERRO: {exc}")
        sys.exit(1)

    print(f"\nResumo:")
    print(f"  Linhas      : {len(df)}")
    print(f"  Municipios  : {df['cod_ibge'].nunique()}")
    anos_disp = sorted(df['ano'].dropna().unique().tolist())
    print(f"  Anos        : {anos_disp}")

    if len(anos_disp) < len(ANOS_DEFAULT):
        faltando = [a for a in ANOS_DEFAULT if a not in anos_disp]
        print(f"  [AVISO] Anos sem dados: {faltando}")

    print(f"\nFontes utilizadas:")
    for fonte, n in df.groupby("fonte").size().items():
        print(f"  {n:>4} obs: {fonte}")

    print("\nConcluido.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Baixa populacao municipal 2018-2025")
    parser.add_argument("--force", action="store_true", help="Forca re-download")
    parser.add_argument("--sem-datasus", action="store_true", help="Pula DATASUS, usa IBGE direto")
    args = parser.parse_args()
    main(force=args.force, sem_datasus=args.sem_datasus)
