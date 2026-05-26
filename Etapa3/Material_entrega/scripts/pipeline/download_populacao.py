"""
Valida o arquivo local de populacao municipal indicado pelo professor.

Uso:
    cd Etapa3/Material_entrega/scripts
    python -m pipeline.download_populacao
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1]))

from pipeline.config import CONFIG
from pipeline.populacao import ANOS_DEFAULT, ARQUIVO_POPULACAO_TABNET, obter_populacao


def main() -> None:
    print("=" * 60)
    print("Validacao da populacao municipal")
    print(f"  UF      : {CONFIG.uf}")
    print(f"  Anos    : {ANOS_DEFAULT[0]} a {ANOS_DEFAULT[-1]}")
    print(f"  Arquivo : {ARQUIVO_POPULACAO_TABNET.name}")
    print("=" * 60)

    try:
        df = obter_populacao(anos=ANOS_DEFAULT, uf=CONFIG.uf)
    except Exception as exc:
        print(f"\nERRO: {exc}")
        sys.exit(1)

    print("\nResumo:")
    print(f"  Linhas      : {len(df)}")
    print(f"  Municipios  : {df['cod_ibge'].nunique()}")
    print(f"  Anos        : {sorted(df['ano'].dropna().unique().tolist())}")

    print("\nFonte utilizada:")
    for fonte, n in df.groupby("fonte").size().items():
        print(f"  {n:>4} obs: {fonte}")

    print("\nConcluido.")


if __name__ == "__main__":
    main()
