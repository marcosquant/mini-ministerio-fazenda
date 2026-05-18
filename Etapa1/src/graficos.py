"""
Geração dos 8 gráficos dos demonstrativos fiscais.
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import warnings
from .config import ANOS, OUTPUT_DIR, CORES

warnings.filterwarnings('ignore')

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'font.size': 10,
    'axes.titlesize': 12,
    'axes.titleweight': 'bold',
    'figure.dpi': 150,
})

_ANOS_LABELS = [str(a) for a in ANOS]


def _fmt_bilhoes(val, pos):
    if abs(val) >= 1:
        return f'R$ {val:.1f} bi'
    return f'R$ {val*1000:.0f} mi'


def _fmt_milhoes(val, pos):
    return f'R$ {val:.0f} mi'


def _salvar(nome_arquivo):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, nome_arquivo)
    plt.savefig(path, bbox_inches='tight', dpi=150)
    plt.close()
    return path


def grafico1_resultado_primario_estado(resultados_estado):
    """Gráfico de barras: Resultado Primário - Estado MS (R$ bi)."""
    fig, ax = plt.subplots(figsize=(10, 5))
    rp = [resultados_estado[a]['Resultado_Primario'] / 1e9 for a in ANOS]
    cores = [CORES['verde_claro'] if v >= 0 else CORES['vermelho'] for v in rp]
    bars = ax.bar(_ANOS_LABELS, rp, color=cores, edgecolor='white', linewidth=0.5, width=0.6)
    ax.axhline(y=0, color='black', linewidth=0.8)
    for bar, val in zip(bars, rp):
        ax.text(bar.get_x() + bar.get_width() / 2,
                bar.get_height() + (0.05 if val >= 0 else -0.15),
                f'R$ {val:.2f} bi', ha='center',
                va='bottom' if val >= 0 else 'top', fontsize=9, fontweight='bold')
    ax.set_title('Resultado Primário - Estado do Mato Grosso do Sul\n2019-2024',
                 fontweight='bold', pad=15)
    ax.set_ylabel('R$ Bilhões', fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_bilhoes))
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    return _salvar('grafico1_resultado_primario_estado.png')


def grafico2_resultado_primario_capital(resultados_capital):
    """Gráfico de barras: Resultado Primário - Campo Grande (R$ mi)."""
    fig, ax = plt.subplots(figsize=(10, 5))
    rp = [resultados_capital[a]['Resultado_Primario'] / 1e6 for a in ANOS]
    cores = [CORES['verde_claro'] if v >= 0 else CORES['vermelho'] for v in rp]
    bars = ax.bar(_ANOS_LABELS, rp, color=cores, edgecolor='white', linewidth=0.5, width=0.6)
    ax.axhline(y=0, color='black', linewidth=0.8)
    for bar, val in zip(bars, rp):
        offset = 5 if val >= 0 else -15
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + offset,
                f'R$ {val:.1f} mi', ha='center',
                va='bottom' if val >= 0 else 'top', fontsize=9, fontweight='bold')
    ax.set_title('Resultado Primário - Campo Grande (MS)\n2019-2024', fontweight='bold', pad=15)
    ax.set_ylabel('R$ Milhões', fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_milhoes))
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    return _salvar('grafico2_resultado_primario_capital.png')


def grafico3_rec_desp_primarias_estado(resultados_estado):
    """Gráfico de barras agrupadas: Receitas vs Despesas Primárias - Estado MS."""
    fig, ax = plt.subplots(figsize=(11, 6))
    rec = [resultados_estado[a]['Rec_Primaria_Total'] / 1e9 for a in ANOS]
    desp = [resultados_estado[a]['Desp_Primaria_Total'] / 1e9 for a in ANOS]
    x = np.arange(len(ANOS))
    w = 0.35
    ax.bar(x - w/2, rec, w, label='Receitas Primárias', color=CORES['azul_claro'], edgecolor='white')
    ax.bar(x + w/2, desp, w, label='Despesas Primárias', color=CORES['laranja'], edgecolor='white')
    ax.set_xticks(x)
    ax.set_xticklabels(_ANOS_LABELS)
    ax.set_title('Receitas vs Despesas Primárias - Estado do MS\n2019-2024', fontweight='bold', pad=15)
    ax.set_ylabel('R$ Bilhões', fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_bilhoes))
    ax.legend(loc='lower right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    return _salvar('grafico3_rec_desp_primarias_estado.png')


def grafico4_rec_desp_primarias_capital(resultados_capital):
    """Gráfico de barras agrupadas: Receitas vs Despesas Primárias - Campo Grande."""
    fig, ax = plt.subplots(figsize=(11, 6))
    rec = [resultados_capital[a]['Rec_Primaria_Total'] / 1e6 for a in ANOS]
    desp = [resultados_capital[a]['Desp_Primaria_Total'] / 1e6 for a in ANOS]
    x = np.arange(len(ANOS))
    w = 0.35
    ax.bar(x - w/2, rec, w, label='Receitas Primárias', color=CORES['azul_claro'], edgecolor='white')
    ax.bar(x + w/2, desp, w, label='Despesas Primárias', color=CORES['laranja'], edgecolor='white')
    ax.set_xticks(x)
    ax.set_xticklabels(_ANOS_LABELS)
    ax.set_title('Receitas vs Despesas Primárias - Campo Grande\n2019-2024', fontweight='bold', pad=15)
    ax.set_ylabel('R$ Milhões', fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_milhoes))
    ax.legend(loc='lower right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    return _salvar('grafico4_rec_desp_primarias_capital.png')


def grafico5_composicao_despesas_estado(resultados_estado):
    """Gráfico de pizza duplo: Composição das Despesas - Estado MS (2019 e 2024)."""
    fig, axes = plt.subplots(1, 2, figsize=(12, 6))
    for ano_ref, ax in zip([2019, 2024], axes):
        r = resultados_estado[ano_ref]
        valores = [r['DC_Pessoal']/1e9, r['DC_Juros']/1e9, r['DC_Outras']/1e9,
                   r['DK_Investimentos']/1e9, r['DK_Demais_Inversoes']/1e9, r['DK_Financeiras']/1e9]
        labels = ['Pessoal', 'Juros/Enc.Dívida', 'Outras Desp.Correntes',
                  'Investimentos', 'Demais Inversões', 'Desp.Fin.Capital']
        cores_pizza = ['#2E75B6', '#C00000', '#70AD47', '#ED7D31', '#A5A5A5', '#FFC000']
        pares = [(v, l, c) for v, l, c in zip(valores, labels, cores_pizza) if v > 0]
        if pares:
            v_f, l_f, c_f = zip(*pares)
            _, _, autotexts = ax.pie(
                v_f, labels=None, colors=c_f, autopct='%1.1f%%', startangle=90,
                pctdistance=0.75, wedgeprops={'edgecolor': 'white', 'linewidth': 1}
            )
            for at in autotexts:
                at.set_fontsize(8)
            ax.legend(l_f, loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=2, fontsize=8)
        ax.set_title(f'Composição das Despesas - Estado MS\n{ano_ref}', fontweight='bold')
    plt.tight_layout()
    return _salvar('grafico5_composicao_despesas_estado.png')


def grafico6_composicao_receitas_estado(resultados_estado):
    """Gráfico de pizza duplo: Composição das Receitas - Estado MS (2019 e 2024)."""
    fig, axes = plt.subplots(1, 2, figsize=(12, 6))
    for ano_ref, ax in zip([2019, 2024], axes):
        r = resultados_estado[ano_ref]
        valores = [
            r['RC_ICMS']/1e9, r['RC_IPVA']/1e9,
            (r['RC_Impostos_Taxas'] - r['RC_ICMS'] - r['RC_IPVA'])/1e9,
            r['RC_FPE']/1e9,
            (r['RC_Transferencias'] - r['RC_FPE'])/1e9,
            r['RC_Financeiras']/1e9, r['RC_Demais']/1e9,
        ]
        labels = ['ICMS', 'IPVA', 'Outros Impostos/Taxas',
                  'Cota-Parte FPE', 'Outras Transferências',
                  'Rec.Financeiras', 'Demais Rec.Correntes']
        cores_pizza = ['#1F4E79', '#2E75B6', '#9DC3E6', '#375623', '#70AD47', '#FFC000', '#ED7D31']
        pares = [(v, l, c) for v, l, c in zip(valores, labels, cores_pizza) if v > 0]
        if pares:
            v_f, l_f, c_f = zip(*pares)
            _, _, autotexts = ax.pie(
                v_f, labels=None, colors=c_f, autopct='%1.1f%%', startangle=90,
                pctdistance=0.75, wedgeprops={'edgecolor': 'white', 'linewidth': 1}
            )
            for at in autotexts:
                at.set_fontsize(8)
            ax.legend(l_f, loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=8)
        ax.set_title(f'Composição das Receitas - Estado MS\n{ano_ref}', fontweight='bold')
    plt.tight_layout()
    return _salvar('grafico6_composicao_receitas_estado.png')


def grafico7_resultado_orcamentario(resultados_estado, resultados_capital):
    """Gráfico duplo: Resultado Orçamentário - Estado e Campo Grande."""
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))

    ro_est = [resultados_estado[a]['Resultado_Orcamentario'] / 1e9 for a in ANOS]
    cores = [CORES['verde_claro'] if v >= 0 else CORES['vermelho'] for v in ro_est]
    axes[0].bar(_ANOS_LABELS, ro_est, color=cores, edgecolor='white', width=0.6)
    axes[0].axhline(y=0, color='black', linewidth=0.8)
    axes[0].set_title('Resultado Orçamentário\nEstado do MS', fontweight='bold')
    axes[0].set_ylabel('R$ Bilhões')
    axes[0].yaxis.set_major_formatter(plt.FuncFormatter(_fmt_bilhoes))
    axes[0].spines['top'].set_visible(False)
    axes[0].spines['right'].set_visible(False)
    axes[0].grid(axis='y', alpha=0.3, linestyle='--')

    ro_cap = [resultados_capital[a]['Resultado_Orcamentario'] / 1e6 for a in ANOS]
    cores = [CORES['verde_claro'] if v >= 0 else CORES['vermelho'] for v in ro_cap]
    axes[1].bar(_ANOS_LABELS, ro_cap, color=cores, edgecolor='white', width=0.6)
    axes[1].axhline(y=0, color='black', linewidth=0.8)
    axes[1].set_title('Resultado Orçamentário\nCampo Grande (MS)', fontweight='bold')
    axes[1].set_ylabel('R$ Milhões')
    axes[1].yaxis.set_major_formatter(plt.FuncFormatter(_fmt_milhoes))
    axes[1].spines['top'].set_visible(False)
    axes[1].spines['right'].set_visible(False)
    axes[1].grid(axis='y', alpha=0.3, linestyle='--')

    plt.tight_layout()
    return _salvar('grafico7_resultado_orcamentario.png')


def grafico8_evolucao_receitas_despesas_estado(resultados_estado):
    """Gráfico de linhas: ICMS, Pessoal e Receita Corrente Total - Estado MS."""
    fig, ax = plt.subplots(figsize=(10, 5))
    icms = [resultados_estado[a]['RC_ICMS'] / 1e9 for a in ANOS]
    pessoal = [resultados_estado[a]['DC_Pessoal'] / 1e9 for a in ANOS]
    rec_total = [resultados_estado[a]['RC_Total'] / 1e9 for a in ANOS]
    ax.plot(_ANOS_LABELS, icms, 'o-', color=CORES['azul'], linewidth=2,
            markersize=7, label='ICMS')
    ax.plot(_ANOS_LABELS, pessoal, 's-', color=CORES['laranja'], linewidth=2,
            markersize=7, label='Pessoal e Enc.Sociais')
    ax.plot(_ANOS_LABELS, rec_total, '^--', color=CORES['verde'], linewidth=2,
            markersize=7, label='Receita Corrente Total')
    ax.set_title('Evolução de Receitas e Despesas Selecionadas\nEstado do MS - 2019-2024',
                 fontweight='bold', pad=15)
    ax.set_ylabel('R$ Bilhões')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_bilhoes))
    ax.legend(loc='upper left')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(alpha=0.3, linestyle='--')
    plt.tight_layout()
    return _salvar('grafico8_evolucao_receitas_despesas_estado.png')


def gerar_todos(resultados_estado, resultados_capital):
    """
    Gera os 8 gráficos e retorna lista de caminhos.

    Parâmetros
    ----------
    resultados_estado : dict  {ano: dict_metricas}
    resultados_capital : dict {ano: dict_metricas}
    """
    return [
        grafico1_resultado_primario_estado(resultados_estado),
        grafico2_resultado_primario_capital(resultados_capital),
        grafico3_rec_desp_primarias_estado(resultados_estado),
        grafico4_rec_desp_primarias_capital(resultados_capital),
        grafico5_composicao_despesas_estado(resultados_estado),
        grafico6_composicao_receitas_estado(resultados_estado),
        grafico7_resultado_orcamentario(resultados_estado, resultados_capital),
        grafico8_evolucao_receitas_despesas_estado(resultados_estado),
    ]
