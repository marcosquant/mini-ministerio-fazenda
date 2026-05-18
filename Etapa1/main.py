"""
Mini Ministério da Fazenda - Etapa 1
Estado: Mato Grosso do Sul (MS) | Capital: Campo Grande

Ponto de entrada do processamento FINBRA 2019-2024.
Execute: python main.py
"""

import warnings
warnings.filterwarnings('ignore')

from src.config import ANOS, OUTPUT_DIR
from src.calculadora import calcular_estado, calcular_capital
from src.demonstrativos import (
    build_demonstrativo1_estado,
    build_demonstrativo1_capital,
    build_demonstrativo2,
)
from src.exportador import exportar_excel, exportar_csv, exportar_base_dados_csv
from src.graficos import gerar_todos


def main():
    print("=" * 60)
    print("PROCESSAMENTO FINBRA - MATO GROSSO DO SUL")
    print("=" * 60)

    # ----------------------------------------------------------------
    # 1. Calcular indicadores por ano
    # ----------------------------------------------------------------
    print("\n[1/4] Processando dados FINBRA...")
    resultados_estado = {}
    resultados_capital = {}

    for ano in ANOS:
        print(f"  Ano {ano}:", end=' ', flush=True)
        resultados_estado[ano] = calcular_estado(ano)
        rp_est = resultados_estado[ano]['Resultado_Primario']
        print(f"Estado R$ {rp_est/1e9:.2f} bi |", end=' ', flush=True)

        resultados_capital[ano] = calcular_capital(ano)
        rp_cap = resultados_capital[ano]['Resultado_Primario']
        print(f"Capital R$ {rp_cap/1e6:.2f} mi")

    # ----------------------------------------------------------------
    # 2. Construir demonstrativos
    # ----------------------------------------------------------------
    print("\n[2/4] Construindo demonstrativos...")
    dem1_estado = build_demonstrativo1_estado(resultados_estado)
    dem1_capital = build_demonstrativo1_capital(resultados_capital)
    dem2_estado = build_demonstrativo2(resultados_estado)
    dem2_capital = build_demonstrativo2(resultados_capital)

    # ----------------------------------------------------------------
    # 3. Exportar Excel e CSV
    # ----------------------------------------------------------------
    print("\n[3/4] Exportando arquivos...")
    excel_path = exportar_excel(dem1_estado, dem1_capital, dem2_estado, dem2_capital)
    print(f"  Excel: {excel_path}")

    csv_paths = exportar_csv(dem1_estado, dem1_capital, dem2_estado, dem2_capital)
    for p in csv_paths:
        print(f"  CSV:   {p}")

    base_csv_paths = exportar_base_dados_csv()
    print(f"  Base de dados: {len(base_csv_paths)} arquivos CSV em output/base_dados_csv/")

    # ----------------------------------------------------------------
    # 4. Gerar gráficos
    # ----------------------------------------------------------------
    print("\n[4/4] Gerando gráficos...")
    graficos = gerar_todos(resultados_estado, resultados_capital)
    for g in graficos:
        print(f"  {g}")

    # ----------------------------------------------------------------
    # Resumo final
    # ----------------------------------------------------------------
    print("\n" + "=" * 60)
    print("RESUMO DOS RESULTADOS")
    print("=" * 60)

    print("\n--- ESTADO DO MATO GROSSO DO SUL ---")
    print(f"{'Ano':<6} {'Rec.Primária':>16} {'Desp.Primária':>16} "
          f"{'Res.Primário':>16} {'Res.Orçam.':>16}")
    print(f"{'':6} {'(R$ bi)':>16} {'(R$ bi)':>16} {'(R$ bi)':>16} {'(R$ bi)':>16}")
    for ano in ANOS:
        r = resultados_estado[ano]
        print(f"{ano:<6} {r['Rec_Primaria_Total']/1e9:>16.2f} "
              f"{r['Desp_Primaria_Total']/1e9:>16.2f} "
              f"{r['Resultado_Primario']/1e9:>16.2f} "
              f"{r['Resultado_Orcamentario']/1e9:>16.2f}")

    print("\n--- CAMPO GRANDE (MS) ---")
    print(f"{'Ano':<6} {'Rec.Primária':>16} {'Desp.Primária':>16} "
          f"{'Res.Primário':>16} {'Res.Orçam.':>16}")
    print(f"{'':6} {'(R$ mi)':>16} {'(R$ mi)':>16} {'(R$ mi)':>16} {'(R$ mi)':>16}")
    for ano in ANOS:
        r = resultados_capital[ano]
        print(f"{ano:<6} {r['Rec_Primaria_Total']/1e6:>16.2f} "
              f"{r['Desp_Primaria_Total']/1e6:>16.2f} "
              f"{r['Resultado_Primario']/1e6:>16.2f} "
              f"{r['Resultado_Orcamentario']/1e6:>16.2f}")

    print(f"\nTodos os arquivos salvos em: {OUTPUT_DIR}")
    print("Processamento concluído!")


if __name__ == '__main__':
    main()
