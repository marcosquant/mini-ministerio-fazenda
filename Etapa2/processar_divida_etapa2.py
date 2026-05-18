"""
Etapa 2 – Diagnóstico do Endividamento de MS e Campo Grande (2019-2024)
Fontes:
  Anexo 2 RGF – Demonstrativo da Dívida Consolidada Líquida (xlsx)
  Anexo 5 RGF – Demonstrativo da Disponibilidade de Caixa e dos Restos a Pagar (csv)

Uso:
  python processar_divida_etapa2.py
"""

from config  import ANOS
from extrair import extrair_dados_uf, extrair_liquidez_uf, extrair_servico_divida_etapa1
from utils   import get, pct, rcl6
from exportar import gerar_quadro_csv, gerar_indicadores_csv, gerar_excel
from config  import LINHAS_ESTADO, LINHAS_CAPITAL


def main():
    # ── Extração ────────────────────────────────────────────────────────────
    print("Lendo Anexo 2 (xlsx)...")
    ms     = extrair_dados_uf("Estados.DF", "MS")
    cg     = extrair_dados_uf("Capitais",   "MS")

    print("Lendo Anexo 5 (csv)...")
    ms_liq = extrair_liquidez_uf("MS")

    print("Lendo servico da divida (Etapa 1)...")
    servico_divida = extrair_servico_divida_etapa1()

    # ── CSVs ────────────────────────────────────────────────────────────────
    print("\nGerando tabelas de decomposição da dívida...")
    gerar_quadro_csv(LINHAS_ESTADO,  ms, "quadro_decomp_divida_MS_estado.csv")
    gerar_quadro_csv(LINHAS_CAPITAL, cg, "quadro_decomp_divida_CG_capital.csv")

    print("Gerando quadro de indicadores...")
    gerar_indicadores_csv(ms, cg, ms_liq, "indicadores_endividamento.csv", servico_divida)

    # ── Sumário no terminal ─────────────────────────────────────────────────
    print()
    print("=" * 70)
    print("SUMÁRIO – DCL / RCL Ajustada | 3º Quadrimestre")
    print("=" * 70)
    print(f"{'Ano':<6} {'DCL MS':>10} {'RCL MS':>10} {'DCL/RCL%':>10} | "
          f"{'DCL CG':>10} {'RCL CG':>10} {'DCL/RCL%':>10}")
    print("-" * 70)
    for ano in ANOS:
        dcl_ms = get(ms[ano], "dcl")
        r_ms   = rcl6(ms[ano])
        dcl_cg = get(cg[ano], "dcl")
        r_cg   = rcl6(cg[ano])
        print(f"{ano:<6} {dcl_ms/1e6:>10.3f} {r_ms/1e6:>10.3f} {pct(dcl_ms,r_ms):>9.2f}% | "
              f"{dcl_cg/1e6:>10.3f} {r_cg/1e6:>10.3f} {pct(dcl_cg,r_cg):>9.2f}%")

    print()
    print("PASSIVOS FORA DA DC + LIQUIDEZ – MS Estado")
    print(f"{'Ano':<6} {'P.Atuarial':>12} {'PA/RCL%':>9} | "
          f"{'Disp.Bruta':>12} {'Disp.Líq(após)':>16} {'DL/RCL%':>9}")
    print("-" * 75)
    for ano in ANOS:
        pa   = get(ms[ano], "passivo_atuarial")
        r    = rcl6(ms[ano])
        db   = ms_liq[ano].get("disp_bruta_a5", 0)
        dl   = ms_liq[ano].get("disp_liq_apos")
        print(
            f"{ano:<6} "
            f"{ f'{pa/1e6:.3f}' if pa else '–':>12} "
            f"{ f'{pct(pa,r):.2f}%' if pa else '–':>9} | "
            f"{ f'{db/1e6:.3f}' if db else '–':>12} "
            f"{ f'{dl/1e6:.3f}' if dl is not None else '–':>16} "
            f"{ f'{pct(dl,r):.2f}%' if dl is not None else '–':>9}"
        )

    # ── Excel ───────────────────────────────────────────────────────────────
    print()
    print("Gerando Excel...")
    try:
        gerar_excel(ms, cg, ms_liq, "Etapa2_Diagnostico_Endividamento_MS_CG.xlsx", servico_divida)
    except Exception as exc:
        import traceback
        print(f"  Aviso: não foi possível gerar Excel – {exc}")
        traceback.print_exc()

    print()
    print("Pronto. Todos os arquivos salvos em Etapa2/Output/")


if __name__ == "__main__":
    main()
