"""Constantes globais, caminhos e estruturas de dados do projeto Etapa 2."""

import os

BASE = os.path.join(os.path.dirname(__file__), "base_dados")
OUT  = os.path.join(os.path.dirname(__file__), "Output")
ANOS = [2019, 2020, 2021, 2022, 2023, 2024]

# ---------------------------------------------------------------------------
# Mapeamento de contas – Anexo 2 RGF
# Normaliza variações de nome entre anos para uma chave canônica.
# ---------------------------------------------------------------------------
CONTA_MAP = {
    # Dívida Consolidada
    "DÍVIDA CONSOLIDADA - DC (I)": "dc",
    "Dívida Contratual": "contratual",
    "Empréstimos": "emprestimos",
    "Internos": "__ambig_internos",   # resolvido por contexto em extrair.py
    "Externos": "__ambig_externos",
    "Reestruturação da Dívida de Estados e Municípios": "reestruturacao",
    "Financiamentos": "financiamentos",
    "Parcelamento e Renegociação de Dívidas": "parcelamento",
    "De Contribuições Previdenciárias": "parcel_prev",
    "De Demais Contribuições Sociais": "parcel_demais",
    "Do FGTS": "parcel_fgts",
    "Com Instituição Não Financeira": "parcel_nao_fin",
    "De Tributos": "parcel_tributos",
    "Precatórios Posteriores a 05/05/2000 (inclusive) Vencidos e Não Pagos": "precatorios",
    "Dívida Mobiliária": "mobiliaria",
    # Deduções
    "DEDUÇÕES (II)": "deducoes",
    "Disponibilidade de Caixa Bruta": "disp_bruta",
    "(-) Restos a Pagar Processados": "rp_processados",
    "Disponibilidade de Caixa": "disp_caixa",
    "Demais Haveres Financeiros": "demais_haveres",
    "(-) Depósitos Restituíveis e Valores Vinculados": "depositos_restitui",
    # Totais
    "DÍVIDA CONSOLIDADA LÍQUIDA (DCL) (III) = (I - II)": "dcl",
    # RCL – variações de nome entre anos
    "RECEITA CORRENTE LÍQUIDA - RCL": "rcl",
    "RECEITA CORRENTE LÍQUIDA - RCL (IV)": "rcl",
    # Transferências obrigatórias (art. 166-A) – variações tipográficas
    "(-) Transferências Obrigatórias da União Relativas às Emendas Individuais (art. 166-A, § 1°, da CF) (V)": "transf_emendas",
    "(-) Transferências Obrigatórias da União Relativas às Emendas Individuais (art. 166-A, § 1º, da CF) (V)": "transf_emendas",
    # RCL Ajustada – variações de nome entre anos
    "= RECEITA CORRENTE LÍQUIDA AJUSTADA PARA CÁLCULO DOS LIMITES DE ENDIVIDAMENTO (VI) = (IV - V)": "rcl_ajustada",
    "= RECEITA CORRENTE LÍQUIDA AJUSTADA PARA PARA CÁLCULO DOS LIMITES DE ENDIVIDAMENTO (VI) = (IV - V)": "rcl_ajustada",
    "RECEITA CORRENTE LÍQUIDA AJUSTADA PARA CÁLCULO DOS LIMITES DE ENDIVIDAMENTO (VI) = (IV - V)": "rcl_ajustada",
    # Passivos fora da dívida consolidada
    "Passivo Atuarial": "passivo_atuarial",
    "RP Não-Processados": "rp_nao_processados",
    "Depósitos e Consignações Sem Contrapartida": "dep_consig",
    "Apropriação de Depósitos Judiciais": "aprops_dep_judiciais",
    # Limites LRF (capturados mas não usados no output)
    "LIMITE DEFINIDO POR RESOLUÇÃO DO SENADO FEDERAL": "limite_senado",
    "LIMITE DE ALERTA (inciso III do § 1° do art. 59 da LRF)": "limite_alerta",
    "LIMITE DE ALERTA (inciso III do § 1º do art. 59 da LRF)": "limite_alerta",
    "% da DC sobre a RCL (I/RCL)": "_pct_dc_rcl",
    "% da DCL sobre a RCL (III/RCL)": "_pct_dcl_rcl",
    "% da DC sobre a RCL AJUSTADA (I/VI)": "_pct_dc_ajust",
    "% da DCL sobre a RCL AJUSTADA (III/VI)": "_pct_dcl_ajust",
}

# Contas de total consolidado no Anexo 5 (o nome variou a partir de 2023)
ANEXO5_TOTAL_CONTAS = {"TOTAL (III) = (I + II)", "TOTAL (IV) = (I + II + III)"}

# ---------------------------------------------------------------------------
# Estrutura das tabelas de decomposição da dívida
# Cada tupla: (rótulo, chave_canônica, é_linha_de_total)
# ---------------------------------------------------------------------------
LINHAS_ESTADO = [
    ("I - DÍVIDA CONSOLIDADA (DC)", "dc", True),
    ("  Dívida Mobiliária", "mobiliaria", False),
    ("  Dívida Contratual", "contratual", False),
    ("    Empréstimos", "emprestimos", False),
    ("      Internos", "emprest_internos", False),
    ("      Externos", "emprest_externos", False),
    ("    Reestruturação da Dívida", "reestruturacao", False),
    ("    Financiamentos", "financiamentos", False),
    ("      Internos", "financ_internos", False),
    ("      Externos", "financ_externos", False),
    ("    Parcelamento e Renegociação", "parcelamento", False),
    ("      Contribuições Previdenciárias", "parcel_prev", False),
    ("      Demais Contribuições Sociais", "parcel_demais", False),
    ("  Precatórios (pós 05/05/2000)", "precatorios", False),
    ("II - DEDUÇÕES", "deducoes", True),
    ("  Disponibilidade de Caixa Bruta", "disp_bruta", False),
    ("  (-) RP Processados", "rp_processados", False),
    ("  = Disponibilidade de Caixa", "disp_caixa", False),
    ("  Demais Haveres Financeiros", "demais_haveres", False),
    ("III - DCL (I - II)", "dcl", True),
    ("IV - RCL", "rcl", True),
    ("V - Transf. Obrig. União (Emendas)", "transf_emendas", False),
    ("VI - RCL Ajustada (IV - V)", "rcl_ajustada", True),
    ("--- PASSIVOS FORA DA DC ---", None, True),
    ("  Passivo Atuarial (RPPS)", "passivo_atuarial", False),
    ("  RP Não-Processados", "rp_nao_processados", False),
    ("  Dep. e Consig. Sem Contrapartida", "dep_consig", False),
    ("  Apropriação de Dep. Judiciais", "aprops_dep_judiciais", False),
]

LINHAS_CAPITAL = [
    ("I - DÍVIDA CONSOLIDADA (DC)", "dc", True),
    ("  Dívida Contratual", "contratual", False),
    ("    Financiamentos", "financiamentos", False),
    ("      Internos", "financ_internos", False),
    ("      Externos", "financ_externos", False),
    ("    Parcelamento e Renegociação", "parcelamento", False),
    ("      Contribuições Previdenciárias", "parcel_prev", False),
    ("      Demais Contribuições Sociais", "parcel_demais", False),
    ("      Do FGTS", "parcel_fgts", False),
    ("      Com Inst. Não Financeira", "parcel_nao_fin", False),
    ("II - DEDUÇÕES", "deducoes", True),
    ("  Disponibilidade de Caixa Bruta", "disp_bruta", False),
    ("  (-) RP Processados", "rp_processados", False),
    ("  = Disponibilidade de Caixa", "disp_caixa", False),
    ("  Demais Haveres Financeiros", "demais_haveres", False),
    ("III - DCL (I - II)", "dcl", True),
    ("IV - RCL", "rcl", True),
    ("V - Transf. Obrig. União (Emendas)", "transf_emendas", False),
    ("VI - RCL Ajustada (IV - V)", "rcl_ajustada", True),
]
