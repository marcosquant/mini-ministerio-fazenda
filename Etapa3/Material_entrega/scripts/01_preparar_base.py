from __future__ import annotations

from pipeline.area import obter_area_municipios
from pipeline.config import CONFIG, PATHS
from pipeline.geo_distance import obter_distancia_capital
from pipeline.populacao import ANOS_DEFAULT, obter_populacao
from pipeline.readers import ler_despesas_siconfi, ler_pib_municipal, ler_receitas_siconfi
from pipeline.transformations import preparar_base_final
from pipeline.utils import imprimir_resumo_base


def main() -> None:
    PATHS.outputs.mkdir(parents=True, exist_ok=True)

    receitas  = ler_receitas_siconfi()
    despesas  = ler_despesas_siconfi()
    pib       = ler_pib_municipal()
    area      = obter_area_municipios()
    distancia = obter_distancia_capital()
    populacao = obter_populacao(anos=ANOS_DEFAULT, uf=CONFIG.uf)

    base = preparar_base_final(receitas, despesas, pib, area, distancia, populacao)

    saida = PATHS.outputs / CONFIG.output_base
    base.to_csv(
        saida,
        index=False,
        sep=CONFIG.csv_sep,
        decimal=CONFIG.csv_decimal,
        encoding=CONFIG.csv_encoding,
    )

    imprimir_resumo_base(base, saida)


if __name__ == "__main__":
    main()
