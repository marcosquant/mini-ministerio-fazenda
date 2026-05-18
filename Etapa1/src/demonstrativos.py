"""
Construção dos DataFrames dos demonstrativos fiscais.
"""

import pandas as pd
from .config import ANOS


def _fv(v):
    """Formata valor em R$ milhões (2 casas decimais)."""
    return round(v / 1e6, 2)


def build_demonstrativo1_estado(resultados):
    """
    Constrói o 1º Demonstrativo do Estado (Resultado Primário).

    Parâmetros
    ----------
    resultados : dict
        {ano: dict_de_metricas} retornado por calculadora.calcular_estado.

    Retorna
    -------
    pd.DataFrame com índice = linhas do demonstrativo, colunas = anos.
    """
    linhas = [
        'RECEITA CORRENTE',
        '  Impostos, Taxas e Contribuições de Melhoria',
        '    ICMS',
        '    IPVA',
        '  Transferências Correntes',
        '    Cota-Parte do FPE',
        '  Receitas Financeiras Correntes',
        '  Demais Receitas Correntes',
        'SALDO (A) - Receitas Primárias Correntes',
        'RECEITA DE CAPITAL',
        '  Receitas Financeiras de Capital',
        '    Operações de Crédito',
        '    Alienação de Bens',
        '    Amortização de Empréstimos',
        '  Transferências de Capital',
        '  Outras Receitas de Capital',
        'SALDO (B) - Receitas Primárias de Capital',
        'RECEITA PRIMÁRIA TOTAL (C = A + B)',
        'DESPESA CORRENTE',
        '  Pessoal e Encargos Sociais',
        '  Juros e Encargos da Dívida',
        '  Outras Despesas Correntes',
        'SALDO (D) - Despesas Primárias Correntes',
        'DESPESA DE CAPITAL',
        '  Investimentos',
        '  Demais Inversões',
        '  Despesas Financeiras de Capital',
        '    Amortização da Dívida',
        'SALDO (E) - Despesas Primárias de Capital',
        'DESPESA PRIMÁRIA TOTAL (F = D + E)',
        'RESULTADO PRIMÁRIO (G = C - F)',
    ]

    dados = {linha: [] for linha in linhas}

    for ano in ANOS:
        r = resultados[ano]
        dados['RECEITA CORRENTE'].append(_fv(r['RC_Total']))
        dados['  Impostos, Taxas e Contribuições de Melhoria'].append(_fv(r['RC_Impostos_Taxas']))
        dados['    ICMS'].append(_fv(r['RC_ICMS']))
        dados['    IPVA'].append(_fv(r['RC_IPVA']))
        dados['  Transferências Correntes'].append(_fv(r['RC_Transferencias']))
        dados['    Cota-Parte do FPE'].append(_fv(r['RC_FPE']))
        dados['  Receitas Financeiras Correntes'].append(_fv(r['RC_Financeiras']))
        dados['  Demais Receitas Correntes'].append(_fv(r['RC_Demais']))
        dados['SALDO (A) - Receitas Primárias Correntes'].append(_fv(r['Saldo_A']))
        dados['RECEITA DE CAPITAL'].append(_fv(r['RK_Total']))
        dados['  Receitas Financeiras de Capital'].append(_fv(r['RK_Financeiras']))
        dados['    Operações de Crédito'].append(_fv(r['RK_Op_Credito']))
        dados['    Alienação de Bens'].append(_fv(r['RK_Alienacao']))
        dados['    Amortização de Empréstimos'].append(_fv(r['RK_Amort_Emp']))
        dados['  Transferências de Capital'].append(_fv(r['RK_Transferencias']))
        dados['  Outras Receitas de Capital'].append(_fv(r['RK_Outras']))
        dados['SALDO (B) - Receitas Primárias de Capital'].append(_fv(r['Saldo_B']))
        dados['RECEITA PRIMÁRIA TOTAL (C = A + B)'].append(_fv(r['Rec_Primaria_Total']))
        dados['DESPESA CORRENTE'].append(_fv(r['DC_Total']))
        dados['  Pessoal e Encargos Sociais'].append(_fv(r['DC_Pessoal']))
        dados['  Juros e Encargos da Dívida'].append(_fv(r['DC_Juros']))
        dados['  Outras Despesas Correntes'].append(_fv(r['DC_Outras']))
        dados['SALDO (D) - Despesas Primárias Correntes'].append(_fv(r['Saldo_D']))
        dados['DESPESA DE CAPITAL'].append(_fv(r['DK_Total']))
        dados['  Investimentos'].append(_fv(r['DK_Investimentos']))
        dados['  Demais Inversões'].append(_fv(r['DK_Demais_Inversoes']))
        dados['  Despesas Financeiras de Capital'].append(_fv(r['DK_Financeiras']))
        dados['    Amortização da Dívida'].append(_fv(r['DK_Amort_Divida']))
        dados['SALDO (E) - Despesas Primárias de Capital'].append(_fv(r['Saldo_E']))
        dados['DESPESA PRIMÁRIA TOTAL (F = D + E)'].append(_fv(r['Desp_Primaria_Total']))
        dados['RESULTADO PRIMÁRIO (G = C - F)'].append(_fv(r['Resultado_Primario']))

    df = pd.DataFrame(dados, index=ANOS).T
    df.index.name = 'Conta'
    df.columns = [str(a) for a in ANOS]
    return df


def build_demonstrativo1_capital(resultados):
    """
    Constrói o 1º Demonstrativo da Capital (Resultado Primário).

    Parâmetros
    ----------
    resultados : dict
        {ano: dict_de_metricas} retornado por calculadora.calcular_capital.

    Retorna
    -------
    pd.DataFrame com índice = linhas do demonstrativo, colunas = anos.
    """
    linhas = [
        'RECEITA CORRENTE',
        '  Impostos, Taxas e Contribuições de Melhoria',
        '    ISS',
        '    IPTU',
        '  Transferências Correntes',
        '    Cota-Parte do FPM',
        '  Receitas Financeiras Correntes',
        '  Demais Receitas Correntes',
        'SALDO (A) - Receitas Primárias Correntes',
        'RECEITA DE CAPITAL',
        '  Receitas Financeiras de Capital',
        '    Operações de Crédito',
        '    Alienação de Bens',
        '    Amortização de Empréstimos',
        '  Transferências de Capital',
        '  Outras Receitas de Capital',
        'SALDO (B) - Receitas Primárias de Capital',
        'RECEITA PRIMÁRIA TOTAL (C = A + B)',
        'DESPESA CORRENTE',
        '  Pessoal e Encargos Sociais',
        '  Juros e Encargos da Dívida',
        '  Outras Despesas Correntes',
        'SALDO (D) - Despesas Primárias Correntes',
        'DESPESA DE CAPITAL',
        '  Investimentos',
        '  Demais Inversões',
        '  Despesas Financeiras de Capital',
        '    Amortização da Dívida',
        'SALDO (E) - Despesas Primárias de Capital',
        'DESPESA PRIMÁRIA TOTAL (F = D + E)',
        'RESULTADO PRIMÁRIO (G = C - F)',
    ]

    dados = {linha: [] for linha in linhas}

    for ano in ANOS:
        r = resultados[ano]
        dados['RECEITA CORRENTE'].append(_fv(r['RC_Total']))
        dados['  Impostos, Taxas e Contribuições de Melhoria'].append(_fv(r['RC_Impostos_Taxas']))
        dados['    ISS'].append(_fv(r['RC_ISS']))
        dados['    IPTU'].append(_fv(r['RC_IPTU']))
        dados['  Transferências Correntes'].append(_fv(r['RC_Transferencias']))
        dados['    Cota-Parte do FPM'].append(_fv(r['RC_FPM']))
        dados['  Receitas Financeiras Correntes'].append(_fv(r['RC_Financeiras']))
        dados['  Demais Receitas Correntes'].append(_fv(r['RC_Demais']))
        dados['SALDO (A) - Receitas Primárias Correntes'].append(_fv(r['Saldo_A']))
        dados['RECEITA DE CAPITAL'].append(_fv(r['RK_Total']))
        dados['  Receitas Financeiras de Capital'].append(_fv(r['RK_Financeiras']))
        dados['    Operações de Crédito'].append(_fv(r['RK_Op_Credito']))
        dados['    Alienação de Bens'].append(_fv(r['RK_Alienacao']))
        dados['    Amortização de Empréstimos'].append(_fv(r['RK_Amort_Emp']))
        dados['  Transferências de Capital'].append(_fv(r['RK_Transferencias']))
        dados['  Outras Receitas de Capital'].append(_fv(r['RK_Outras']))
        dados['SALDO (B) - Receitas Primárias de Capital'].append(_fv(r['Saldo_B']))
        dados['RECEITA PRIMÁRIA TOTAL (C = A + B)'].append(_fv(r['Rec_Primaria_Total']))
        dados['DESPESA CORRENTE'].append(_fv(r['DC_Total']))
        dados['  Pessoal e Encargos Sociais'].append(_fv(r['DC_Pessoal']))
        dados['  Juros e Encargos da Dívida'].append(_fv(r['DC_Juros']))
        dados['  Outras Despesas Correntes'].append(_fv(r['DC_Outras']))
        dados['SALDO (D) - Despesas Primárias Correntes'].append(_fv(r['Saldo_D']))
        dados['DESPESA DE CAPITAL'].append(_fv(r['DK_Total']))
        dados['  Investimentos'].append(_fv(r['DK_Investimentos']))
        dados['  Demais Inversões'].append(_fv(r['DK_Demais_Inversoes']))
        dados['  Despesas Financeiras de Capital'].append(_fv(r['DK_Financeiras']))
        dados['    Amortização da Dívida'].append(_fv(r['DK_Amort_Divida']))
        dados['SALDO (E) - Despesas Primárias de Capital'].append(_fv(r['Saldo_E']))
        dados['DESPESA PRIMÁRIA TOTAL (F = D + E)'].append(_fv(r['Desp_Primaria_Total']))
        dados['RESULTADO PRIMÁRIO (G = C - F)'].append(_fv(r['Resultado_Primario']))

    df = pd.DataFrame(dados, index=ANOS).T
    df.index.name = 'Conta'
    df.columns = [str(a) for a in ANOS]
    return df


def build_demonstrativo2(resultados):
    """
    Constrói o 2º Demonstrativo (Resultado Orçamentário Consolidado).

    Fórmula: Resultado Orçamentário = Resultado Primário
             + Receitas Financeiras − Despesas Financeiras

    Parâmetros
    ----------
    resultados : dict
        {ano: dict_de_metricas} do estado ou capital.

    Retorna
    -------
    pd.DataFrame com índice = linhas do demonstrativo, colunas = anos.
    """
    linhas = [
        'RECEITAS PRIMÁRIAS',
        'DESPESAS PRIMÁRIAS',
        'RESULTADO PRIMÁRIO',
        'RECEITAS FINANCEIRAS',
        '  Receitas Financeiras Correntes (Valores Mobiliários)',
        '  Operações de Crédito',
        '  Alienação de Bens',
        '  Amortização de Empréstimos (Recebidas)',
        'DESPESAS FINANCEIRAS',
        '  Juros e Encargos da Dívida',
        '  Amortização da Dívida',
        'RESULTADO ORÇAMENTÁRIO',
    ]

    dados = {linha: [] for linha in linhas}

    for ano in ANOS:
        r = resultados[ano]
        dados['RECEITAS PRIMÁRIAS'].append(_fv(r['Rec_Primaria_Total']))
        dados['DESPESAS PRIMÁRIAS'].append(_fv(r['Desp_Primaria_Total']))
        dados['RESULTADO PRIMÁRIO'].append(_fv(r['Resultado_Primario']))
        dados['RECEITAS FINANCEIRAS'].append(_fv(r['Rec_Financeiras_Total']))
        dados['  Receitas Financeiras Correntes (Valores Mobiliários)'].append(_fv(r['RC_Financeiras']))
        dados['  Operações de Crédito'].append(_fv(r['RK_Op_Credito']))
        dados['  Alienação de Bens'].append(_fv(r['RK_Alienacao']))
        dados['  Amortização de Empréstimos (Recebidas)'].append(_fv(r['RK_Amort_Emp']))
        dados['DESPESAS FINANCEIRAS'].append(_fv(r['Desp_Financeiras_Total']))
        dados['  Juros e Encargos da Dívida'].append(_fv(r['DC_Juros']))
        dados['  Amortização da Dívida'].append(_fv(r['DK_Amort_Divida']))
        dados['RESULTADO ORÇAMENTÁRIO'].append(_fv(r['Resultado_Orcamentario']))

    df = pd.DataFrame(dados, index=ANOS).T
    df.index.name = 'Conta'
    df.columns = [str(a) for a in ANOS]
    return df
