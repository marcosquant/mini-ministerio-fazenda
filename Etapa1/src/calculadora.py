"""
Cálculos fiscais: extrai e agrega indicadores por ano para estado e capital.
"""

from .carregador import ler_finbra, get_valor, get_valor_liquido
from .config import UF


def calcular_estado(ano):
    """
    Processa dados do Estado (MS) para um ano.

    Retorna dict com todas as métricas fiscais necessárias para os demonstrativos.
    Inclui Resultado Primário, Resultado Orçamentário e composições.

    Mudança de códigos em 2022: conta_pattern aceita lista de fallbacks.
    """
    rec_est = ler_finbra('ESTDF', 'Receitas', ano, UF)
    desp_est = ler_finbra('ESTDF', 'Despesas', ano, UF)
    desp_emp = desp_est[desp_est['Coluna'] == 'Despesas Empenhadas']

    # --- RECEITAS CORRENTES ---
    rc_total = get_valor_liquido(rec_est, '1.0.0.0.00.0.0')
    rc_impostos = get_valor_liquido(rec_est, '1.1.0.0.00.0.0')
    # IPVA: 1.1.1.8.01.2.0 (até 2021) / 1.1.1.2.51.0.0 (2022+)
    rc_ipva = get_valor_liquido(rec_est, ['1.1.1.8.01.2.0', '1.1.1.2.51.0.0'])
    # ICMS: 1.1.1.8.02.{1,2}.0 (até 2021) / 1.1.1.4.50.{1,2}.0 (2022+)
    rc_icms = (get_valor_liquido(rec_est, ['1.1.1.8.02.1.0', '1.1.1.4.50.1.0']) +
               get_valor_liquido(rec_est, ['1.1.1.8.02.2.0', '1.1.1.4.50.2.0']))
    rc_transferencias = get_valor_liquido(rec_est, '1.7.0.0.00.0.0')
    # FPE: 1.7.1.8.01.1.0 (até 2021) / 1.7.1.1.50.0.0 (2022+)
    rc_fpe = get_valor_liquido(rec_est, ['1.7.1.8.01.1.0', '1.7.1.1.50.0.0'])
    rc_financeiras = get_valor_liquido(rec_est, '1.3.2.0.00.0.0')
    rc_contribuicoes = get_valor_liquido(rec_est, '1.2.0.0.00.0.0')
    rc_patrimonial_total = get_valor_liquido(rec_est, '1.3.0.0.00.0.0')
    rc_servicos = get_valor_liquido(rec_est, '1.6.0.0.00.0.0')
    rc_outras = get_valor_liquido(rec_est, '1.9.0.0.00.0.0')
    rc_demais = (rc_contribuicoes + (rc_patrimonial_total - rc_financeiras)
                 + rc_servicos + rc_outras)
    saldo_A = rc_total - rc_financeiras

    # --- RECEITAS DE CAPITAL ---
    rk_total = get_valor_liquido(rec_est, '2.0.0.0.00.0.0')
    rk_op_credito = get_valor_liquido(rec_est, '2.1.0.0.00.0.0')
    rk_alienacao = get_valor_liquido(rec_est, '2.2.0.0.00.0.0')
    rk_amort_emp = get_valor_liquido(rec_est, '2.3.0.0.00.0.0')
    rk_financeiras = rk_op_credito + rk_alienacao + rk_amort_emp
    rk_transferencias = get_valor_liquido(rec_est, '2.4.0.0.00.0.0')
    rk_outras = rk_total - rk_op_credito - rk_alienacao - rk_amort_emp - rk_transferencias
    saldo_B = rk_total - rk_financeiras
    rec_primaria_total = saldo_A + saldo_B

    # --- DESPESAS CORRENTES ---
    dc_total = get_valor(desp_emp, 'Despesas Empenhadas', '3.0.00.00.00')
    dc_pessoal = get_valor(desp_emp, 'Despesas Empenhadas', '3.1.00.00.00')
    dc_juros = get_valor(desp_emp, 'Despesas Empenhadas', '3.2.00.00.00')
    dc_outras = get_valor(desp_emp, 'Despesas Empenhadas', '3.3.00.00.00')
    saldo_D = dc_total - dc_juros

    # --- DESPESAS DE CAPITAL ---
    dk_total = get_valor(desp_emp, 'Despesas Empenhadas', '4.0.00.00.00')
    dk_investimentos = get_valor(desp_emp, 'Despesas Empenhadas', '4.4.00.00.00')
    dk_aquisicao_credito = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.63.00')
    dk_aquisicao_capital = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.64.00')
    dk_concessao_emp = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.66.00')
    dk_amort_divida = get_valor(desp_emp, 'Despesas Empenhadas', '4.6.00.00.00')
    dk_financeiras = (dk_aquisicao_credito + dk_aquisicao_capital
                      + dk_concessao_emp + dk_amort_divida)
    dk_demais_inversoes = dk_total - dk_investimentos - dk_financeiras
    saldo_E = dk_total - dk_financeiras
    desp_primaria_total = saldo_D + saldo_E
    resultado_primario = rec_primaria_total - desp_primaria_total

    rec_financeiras_total = rc_financeiras + rk_financeiras
    desp_financeiras_total = dc_juros + dk_amort_divida
    resultado_orcamentario = resultado_primario - (desp_financeiras_total - rec_financeiras_total)

    return {
        'RC_Total': rc_total,
        'RC_Impostos_Taxas': rc_impostos,
        'RC_IPVA': rc_ipva,
        'RC_ICMS': rc_icms,
        'RC_Transferencias': rc_transferencias,
        'RC_FPE': rc_fpe,
        'RC_Financeiras': rc_financeiras,
        'RC_Demais': rc_demais,
        'Saldo_A': saldo_A,
        'RK_Total': rk_total,
        'RK_Op_Credito': rk_op_credito,
        'RK_Alienacao': rk_alienacao,
        'RK_Amort_Emp': rk_amort_emp,
        'RK_Financeiras': rk_financeiras,
        'RK_Transferencias': rk_transferencias,
        'RK_Outras': rk_outras,
        'Saldo_B': saldo_B,
        'Rec_Primaria_Total': rec_primaria_total,
        'DC_Total': dc_total,
        'DC_Pessoal': dc_pessoal,
        'DC_Juros': dc_juros,
        'DC_Outras': dc_outras,
        'Saldo_D': saldo_D,
        'DK_Total': dk_total,
        'DK_Investimentos': dk_investimentos,
        'DK_Demais_Inversoes': dk_demais_inversoes,
        'DK_Financeiras': dk_financeiras,
        'DK_Aq_Credito': dk_aquisicao_credito,
        'DK_Aq_Capital': dk_aquisicao_capital,
        'DK_Concessao_Emp': dk_concessao_emp,
        'DK_Amort_Divida': dk_amort_divida,
        'Saldo_E': saldo_E,
        'Desp_Primaria_Total': desp_primaria_total,
        'Resultado_Primario': resultado_primario,
        'Rec_Financeiras_Total': rec_financeiras_total,
        'Desp_Financeiras_Total': desp_financeiras_total,
        'Resultado_Orcamentario': resultado_orcamentario,
    }


def calcular_capital(ano):
    """
    Processa dados da Capital (Campo Grande) para um ano.

    Retorna dict com todas as métricas fiscais necessárias para os demonstrativos.
    Mudança de códigos em 2022: conta_pattern aceita lista de fallbacks.
    """
    rec_cap = ler_finbra('CAP', 'Receitas', ano, UF)
    desp_cap = ler_finbra('CAP', 'Despesas', ano, UF)
    desp_emp = desp_cap[desp_cap['Coluna'] == 'Despesas Empenhadas']

    # --- RECEITAS CORRENTES ---
    cc_total = get_valor_liquido(rec_cap, '1.0.0.0.00.0.0')
    cc_impostos = get_valor_liquido(rec_cap, '1.1.0.0.00.0.0')
    # IPTU: 1.1.1.8.01.1.0 (até 2021) / 1.1.1.2.50.0.0 (2022+)
    cc_iptu = get_valor_liquido(rec_cap, ['1.1.1.8.01.1.0', '1.1.1.2.50.0.0'])
    # ISS: 1.1.1.8.02.3.0 (até 2021) / 1.1.1.4.51.1.0 (2022+)
    cc_iss = get_valor_liquido(rec_cap, ['1.1.1.8.02.3.0', '1.1.1.4.51.1.0'])
    cc_transferencias = get_valor_liquido(rec_cap, '1.7.0.0.00.0.0')
    # FPM: 1.7.1.8.01.2.0 (até 2021) / 1.7.1.1.51.0.0 (2022+)
    cc_fpm = get_valor_liquido(rec_cap, ['1.7.1.8.01.2.0', '1.7.1.1.51.0.0'])
    cc_financeiras = get_valor_liquido(rec_cap, '1.3.2.0.00.0.0')
    cc_contribuicoes = get_valor_liquido(rec_cap, '1.2.0.0.00.0.0')
    cc_patrimonial_total = get_valor_liquido(rec_cap, '1.3.0.0.00.0.0')
    cc_servicos = get_valor_liquido(rec_cap, '1.6.0.0.00.0.0')
    cc_outras = get_valor_liquido(rec_cap, '1.9.0.0.00.0.0')
    cc_demais = (cc_contribuicoes + (cc_patrimonial_total - cc_financeiras)
                 + cc_servicos + cc_outras)
    saldo_A = cc_total - cc_financeiras

    # --- RECEITAS DE CAPITAL ---
    ck_total = get_valor_liquido(rec_cap, '2.0.0.0.00.0.0')
    ck_op_credito = get_valor_liquido(rec_cap, '2.1.0.0.00.0.0')
    ck_alienacao = get_valor_liquido(rec_cap, '2.2.0.0.00.0.0')
    ck_amort_emp = get_valor_liquido(rec_cap, '2.3.0.0.00.0.0')
    ck_financeiras = ck_op_credito + ck_alienacao + ck_amort_emp
    ck_transferencias = get_valor_liquido(rec_cap, '2.4.0.0.00.0.0')
    ck_outras = ck_total - ck_op_credito - ck_alienacao - ck_amort_emp - ck_transferencias
    saldo_B = ck_total - ck_financeiras
    rec_primaria_total = saldo_A + saldo_B

    # --- DESPESAS CORRENTES ---
    dcc_total = get_valor(desp_emp, 'Despesas Empenhadas', '3.0.00.00.00')
    dcc_pessoal = get_valor(desp_emp, 'Despesas Empenhadas', '3.1.00.00.00')
    dcc_juros = get_valor(desp_emp, 'Despesas Empenhadas', '3.2.00.00.00')
    dcc_outras = get_valor(desp_emp, 'Despesas Empenhadas', '3.3.00.00.00')
    saldo_D = dcc_total - dcc_juros

    # --- DESPESAS DE CAPITAL ---
    dkc_total = get_valor(desp_emp, 'Despesas Empenhadas', '4.0.00.00.00')
    dkc_investimentos = get_valor(desp_emp, 'Despesas Empenhadas', '4.4.00.00.00')
    dkc_aq_credito = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.63.00')
    dkc_aq_capital = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.64.00')
    dkc_concessao_emp = get_valor(desp_emp, 'Despesas Empenhadas', '4.5.90.66.00')
    dkc_amort_divida = get_valor(desp_emp, 'Despesas Empenhadas', '4.6.00.00.00')
    dkc_financeiras = (dkc_aq_credito + dkc_aq_capital
                       + dkc_concessao_emp + dkc_amort_divida)
    dkc_demais_inversoes = dkc_total - dkc_investimentos - dkc_financeiras
    saldo_E = dkc_total - dkc_financeiras
    desp_primaria_total = saldo_D + saldo_E
    resultado_primario = rec_primaria_total - desp_primaria_total

    rec_financeiras_total = cc_financeiras + ck_financeiras
    desp_financeiras_total = dcc_juros + dkc_amort_divida
    resultado_orcamentario = resultado_primario - (desp_financeiras_total - rec_financeiras_total)

    return {
        'RC_Total': cc_total,
        'RC_Impostos_Taxas': cc_impostos,
        'RC_IPTU': cc_iptu,
        'RC_ISS': cc_iss,
        'RC_Transferencias': cc_transferencias,
        'RC_FPM': cc_fpm,
        'RC_Financeiras': cc_financeiras,
        'RC_Demais': cc_demais,
        'Saldo_A': saldo_A,
        'RK_Total': ck_total,
        'RK_Op_Credito': ck_op_credito,
        'RK_Alienacao': ck_alienacao,
        'RK_Amort_Emp': ck_amort_emp,
        'RK_Financeiras': ck_financeiras,
        'RK_Transferencias': ck_transferencias,
        'RK_Outras': ck_outras,
        'Saldo_B': saldo_B,
        'Rec_Primaria_Total': rec_primaria_total,
        'DC_Total': dcc_total,
        'DC_Pessoal': dcc_pessoal,
        'DC_Juros': dcc_juros,
        'DC_Outras': dcc_outras,
        'Saldo_D': saldo_D,
        'DK_Total': dkc_total,
        'DK_Investimentos': dkc_investimentos,
        'DK_Demais_Inversoes': dkc_demais_inversoes,
        'DK_Financeiras': dkc_financeiras,
        'DK_Aq_Credito': dkc_aq_credito,
        'DK_Aq_Capital': dkc_aq_capital,
        'DK_Concessao_Emp': dkc_concessao_emp,
        'DK_Amort_Divida': dkc_amort_divida,
        'Saldo_E': saldo_E,
        'Desp_Primaria_Total': desp_primaria_total,
        'Resultado_Primario': resultado_primario,
        'Rec_Financeiras_Total': rec_financeiras_total,
        'Desp_Financeiras_Total': desp_financeiras_total,
        'Resultado_Orcamentario': resultado_orcamentario,
    }
