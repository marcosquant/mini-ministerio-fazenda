"""
Funções de leitura e extração de dados dos arquivos FINBRA.
"""

import zipfile
import io
import pandas as pd
from .config import DATA_DIR


def ler_finbra(tipo_escopo, tipo_tabela, ano, filtro_uf=None):
    """
    Lê arquivo FINBRA e retorna DataFrame filtrado.

    Parâmetros
    ----------
    tipo_escopo : str
        'ESTDF' para estados ou 'CAP' para capitais.
    tipo_tabela : str
        'Receitas' ou 'Despesas'.
    ano : int
        Ano de referência (2019-2024).
    filtro_uf : str, opcional
        Sigla da UF para filtrar (ex: 'MS').

    Retorna
    -------
    pd.DataFrame com colunas padronizadas e valores numéricos.
    """
    if tipo_tabela == 'Receitas':
        nome = f'finbra_{tipo_escopo}_ReceitasOrcamentarias(AnexoI-C)_{ano}.zip'
    else:
        nome = f'finbra_{tipo_escopo}_DespesasOrcamentarias(AnexoI-D)_{ano}.zip'

    caminho = f'{DATA_DIR}/{nome}'

    with zipfile.ZipFile(caminho) as z:
        with z.open('finbra.csv') as f:
            content = f.read().decode('latin-1')

    lines = content.split('\n')
    # Linha 4 (índice 3) = cabeçalho das colunas; primeiras 3 são metadados
    header = lines[3].strip()
    data_lines = [header] + [l.strip() for l in lines[4:] if l.strip()]

    df = pd.read_csv(
        io.StringIO('\n'.join(data_lines)),
        sep=';',
        dtype=str,
        quotechar='"'
    )

    df.columns = ['Instituicao', 'CodIBGE', 'UF', 'Populacao',
                  'Coluna', 'Conta', 'Identificador', 'Valor']

    if filtro_uf:
        df = df[df['UF'] == filtro_uf].copy()

    df['Valor'] = df['Valor'].str.replace(',', '.', regex=False)
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0)
    df['Ano'] = ano
    df['Conta'] = df['Conta'].str.strip().str.strip('"')

    return df


def get_valor(df, coluna_filtro, conta_pattern, soma_pattern=None):
    """
    Extrai o valor de uma conta específica do DataFrame.

    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame retornado por ler_finbra.
    coluna_filtro : str
        Valor da coluna 'Coluna' a filtrar (ex: 'Despesas Empenhadas').
    conta_pattern : str ou list
        Prefixo da conta a buscar. Lista = fallback por mudança de código em 2022.
    soma_pattern : list, opcional
        Lista de padrões para somar (ignora conta_pattern quando fornecido).

    Retorna
    -------
    float
    """
    df_filtrado = df[df['Coluna'] == coluna_filtro]

    if soma_pattern:
        total = 0
        for pat in soma_pattern:
            mask = df_filtrado['Conta'].str.startswith(pat)
            total += df_filtrado.loc[mask, 'Valor'].sum()
        return total

    patterns = conta_pattern if isinstance(conta_pattern, list) else [conta_pattern]
    for pat in patterns:
        mask = df_filtrado['Conta'].str.startswith(pat)
        vals = df_filtrado.loc[mask, 'Valor']
        if len(vals) > 0 and vals.iloc[0] != 0:
            return vals.iloc[0]
    return 0


def get_valor_liquido(df, conta_pattern):
    """
    Calcula receita líquida = Brutas − FUNDEB − Transferências Constitucionais
    − Outras Deduções da Receita.

    A coluna 'Outras Deduções da Receita' foi introduzida em 2022 pela STN para
    registrar a renúncia fiscal decorrente da isenção de ICMS sobre combustíveis
    (Lei Complementar 194/2022). Para anos anteriores a 2022 o valor é zero,
    portanto a subtração não afeta a série histórica.

    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame de receitas retornado por ler_finbra.
    conta_pattern : str ou list
        Prefixo(s) da conta.

    Retorna
    -------
    float
    """
    brutas = get_valor(df, 'Receitas Brutas Realizadas', conta_pattern)
    fundeb = get_valor(df, 'Deduções - FUNDEB', conta_pattern)
    transf = get_valor(df, 'Deduções - Transferências Constitucionais', conta_pattern)
    outras = get_valor(df, 'Outras Deduções da Receita', conta_pattern)
    return brutas - fundeb - transf - outras
