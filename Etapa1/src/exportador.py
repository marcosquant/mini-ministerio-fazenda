"""
Exportação dos demonstrativos para Excel (formatado) e CSV.
"""

import io
import os
import zipfile
import pandas as pd
from .config import ANOS, DATA_DIR, OUTPUT_DIR


def escrever_planilha_formatada(writer, df, nome_aba, titulo, subtitulo=''):
    """
    Escreve uma aba formatada no arquivo Excel.

    Aplica cores por categoria de linha:
    - Seções principais: fundo azul escuro / cinza claro
    - Saldos (A, B, D, E): fundo verde claro
    - Resultados (Primário, Orçamentário): verde (positivo) / vermelho (negativo)
    - Subcontas com 2 espaços: recuo nível 1
    - Subcontas com 4 espaços: recuo nível 2 em itálico

    Parâmetros
    ----------
    writer : pd.ExcelWriter
        Escritor com engine='xlsxwriter'.
    df : pd.DataFrame
        Demonstrativo com índice = linhas, colunas = anos (str).
    nome_aba : str
        Nome da aba no Excel.
    titulo : str
        Título principal (linha 1).
    subtitulo : str, opcional
        Subtítulo (linha 2, se fornecido).
    """
    ws = writer.book.add_worksheet(nome_aba)
    writer.sheets[nome_aba] = ws
    wb = writer.book

    fmt_titulo = wb.add_format({
        'bold': True, 'font_size': 14, 'align': 'center',
        'valign': 'vcenter', 'bg_color': '#1F4E79', 'font_color': 'white',
        'border': 1
    })
    fmt_subtitulo = wb.add_format({
        'bold': True, 'font_size': 11, 'align': 'center',
        'valign': 'vcenter', 'bg_color': '#2E75B6', 'font_color': 'white',
        'border': 1
    })
    fmt_header = wb.add_format({
        'bold': True, 'font_size': 10, 'align': 'center',
        'valign': 'vcenter', 'bg_color': '#BDD7EE', 'border': 1
    })
    fmt_secao = wb.add_format({
        'bold': True, 'font_size': 10, 'bg_color': '#D6E4F0',
        'border': 1, 'indent': 0
    })
    fmt_saldo = wb.add_format({
        'bold': True, 'font_size': 10, 'bg_color': '#E2EFDA',
        'border': 1, 'num_format': '#,##0.00', 'indent': 0
    })
    fmt_resultado_pos = wb.add_format({
        'bold': True, 'font_size': 10, 'bg_color': '#C6EFCE',
        'font_color': '#375623', 'border': 1, 'num_format': '#,##0.00'
    })
    fmt_resultado_neg = wb.add_format({
        'bold': True, 'font_size': 10, 'bg_color': '#FFC7CE',
        'font_color': '#9C0006', 'border': 1, 'num_format': '#,##0.00'
    })
    fmt_subconta = wb.add_format({
        'font_size': 10, 'border': 1, 'num_format': '#,##0.00', 'indent': 1
    })
    fmt_subsubconta = wb.add_format({
        'font_size': 10, 'border': 1, 'num_format': '#,##0.00', 'indent': 2,
        'italic': True
    })
    fmt_numero = wb.add_format({
        'font_size': 10, 'border': 1, 'num_format': '#,##0.00'
    })
    fmt_nota = wb.add_format({
        'font_size': 8, 'italic': True, 'font_color': '#7F7F7F'
    })

    ws.merge_range(0, 0, 0, len(ANOS), titulo, fmt_titulo)
    ws.set_row(0, 30)

    if subtitulo:
        ws.merge_range(1, 0, 1, len(ANOS), subtitulo, fmt_subtitulo)
        ws.set_row(1, 20)
        row_start = 2
    else:
        row_start = 1

    ws.merge_range(row_start, 0, row_start, len(ANOS), 'Valores em R$ milhões', fmt_nota)
    row_start += 1

    ws.write(row_start, 0, 'CONTA / PERÍODO', fmt_header)
    for i, ano in enumerate(ANOS):
        ws.write(row_start, i + 1, str(ano), fmt_header)
    ws.set_row(row_start, 20)
    row_start += 1

    secoes_principais = {
        'RECEITA CORRENTE', 'RECEITA DE CAPITAL', 'RECEITA PRIMÁRIA TOTAL (C = A + B)',
        'DESPESA CORRENTE', 'DESPESA DE CAPITAL', 'DESPESA PRIMÁRIA TOTAL (F = D + E)',
        'RESULTADO PRIMÁRIO (G = C - F)',
        'RECEITAS PRIMÁRIAS', 'DESPESAS PRIMÁRIAS', 'RESULTADO PRIMÁRIO',
        'RECEITAS FINANCEIRAS', 'DESPESAS FINANCEIRAS', 'RESULTADO ORÇAMENTÁRIO',
    }
    saldos = {
        'SALDO (A) - Receitas Primárias Correntes',
        'SALDO (B) - Receitas Primárias de Capital',
        'SALDO (D) - Despesas Primárias Correntes',
        'SALDO (E) - Despesas Primárias de Capital',
    }
    resultados_linhas = {
        'RECEITA PRIMÁRIA TOTAL (C = A + B)',
        'DESPESA PRIMÁRIA TOTAL (F = D + E)',
        'RESULTADO PRIMÁRIO (G = C - F)',
        'RESULTADO PRIMÁRIO',
        'RESULTADO ORÇAMENTÁRIO',
    }

    for idx_linha, (conta, row_data) in enumerate(df.iterrows()):
        row = row_start + idx_linha
        conta_stripped = conta.strip()

        if conta_stripped in resultados_linhas:
            ws.write(row, 0, conta_stripped, fmt_secao)
            for i, v in enumerate(row_data):
                fmt_val = fmt_resultado_pos if v >= 0 else fmt_resultado_neg
                ws.write_number(row, i + 1, v, fmt_val)
        elif conta_stripped in secoes_principais:
            ws.write(row, 0, conta_stripped, fmt_secao)
            for i, v in enumerate(row_data):
                ws.write_number(row, i + 1, v, fmt_numero)
        elif conta_stripped in saldos:
            ws.write(row, 0, conta_stripped, fmt_secao)
            for i, v in enumerate(row_data):
                ws.write_number(row, i + 1, v, fmt_saldo)
        elif conta.startswith('    '):
            ws.write(row, 0, conta_stripped, fmt_subsubconta)
            for i, v in enumerate(row_data):
                ws.write_number(row, i + 1, v, fmt_subsubconta)
        elif conta.startswith('  '):
            ws.write(row, 0, conta_stripped, fmt_subconta)
            for i, v in enumerate(row_data):
                ws.write_number(row, i + 1, v, fmt_subconta)
        else:
            ws.write(row, 0, conta_stripped, fmt_secao)
            for i, v in enumerate(row_data):
                ws.write_number(row, i + 1, v, fmt_numero)

        ws.set_row(row, 18)

    ws.set_column(0, 0, 55)
    for i in range(len(ANOS)):
        ws.set_column(i + 1, i + 1, 14)


def exportar_excel(dem1_estado, dem1_capital, dem2_estado, dem2_capital):
    """
    Gera o arquivo Excel com 4 abas formatadas.

    Parâmetros
    ----------
    dem1_estado, dem1_capital, dem2_estado, dem2_capital : pd.DataFrame

    Retorna
    -------
    str : caminho do arquivo gerado.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    excel_path = os.path.join(OUTPUT_DIR, 'demonstrativos_fiscais_MS_2019_2024.xlsx')

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        escrever_planilha_formatada(
            writer, dem1_estado, 'Dem1_Estado_MS',
            'PRIMEIRO DEMONSTRATIVO - ESTADO DO MATO GROSSO DO SUL',
            'Resultado Primário e Orçamentário - 2019 a 2024'
        )
        escrever_planilha_formatada(
            writer, dem1_capital, 'Dem1_Capital_CG',
            'PRIMEIRO DEMONSTRATIVO - CAMPO GRANDE (MS)',
            'Resultado Primário e Orçamentário - 2019 a 2024'
        )
        escrever_planilha_formatada(
            writer, dem2_estado, 'Dem2_Estado_MS',
            'SEGUNDO DEMONSTRATIVO - ESTADO DO MATO GROSSO DO SUL',
            'Consolidação do Resultado Orçamentário - 2019 a 2024'
        )
        escrever_planilha_formatada(
            writer, dem2_capital, 'Dem2_Capital_CG',
            'SEGUNDO DEMONSTRATIVO - CAMPO GRANDE (MS)',
            'Consolidação do Resultado Orçamentário - 2019 a 2024'
        )

    return excel_path


def exportar_base_dados_csv(filtro_uf='MS'):
    """
    Consolida os arquivos FINBRA (base_dados/*.zip) em 4 CSVs filtrados por UF,
    um por combinação de escopo (ESTDF/CAP) e tabela (Receitas/Despesas),
    com todos os anos empilhados e coluna 'Ano' adicionada.

    Saída em output/base_dados_csv/:
        base_ESTDF_Receitas_MS.csv
        base_ESTDF_Despesas_MS.csv
        base_CAP_Receitas_MS.csv
        base_CAP_Despesas_MS.csv

    Parâmetros
    ----------
    filtro_uf : str
        Sigla da UF para filtrar (padrão 'MS').

    Retorna
    -------
    list[str] : caminhos dos 4 arquivos CSV gerados.
    """
    dest = os.path.join(OUTPUT_DIR, 'base_dados_csv')
    os.makedirs(dest, exist_ok=True)

    grupos = {
        ('ESTDF', 'Receitas'): [],
        ('ESTDF', 'Despesas'): [],
        ('CAP',   'Receitas'): [],
        ('CAP',   'Despesas'): [],
    }

    for nome_zip in sorted(os.listdir(DATA_DIR)):
        if not nome_zip.endswith('.zip'):
            continue

        # Identifica escopo e tabela pelo nome do arquivo
        if 'ESTDF' in nome_zip:
            escopo = 'ESTDF'
        elif 'CAP' in nome_zip:
            escopo = 'CAP'
        else:
            continue

        if 'Receitas' in nome_zip:
            tabela = 'Receitas'
        elif 'Despesas' in nome_zip:
            tabela = 'Despesas'
        else:
            continue

        # Extrai ano do nome (últimos 4 dígitos antes de .zip)
        ano = int(nome_zip[-8:-4])

        caminho_zip = os.path.join(DATA_DIR, nome_zip)
        with zipfile.ZipFile(caminho_zip) as z:
            with z.open('finbra.csv') as f:
                content = f.read().decode('latin-1')

        lines = content.split('\n')
        header = lines[3].strip()
        data_lines = [header] + [l.strip() for l in lines[4:] if l.strip()]

        df = pd.read_csv(
            io.StringIO('\n'.join(data_lines)),
            sep=';',
            dtype=str,
            quotechar='"'
        )

        # Filtra UF e adiciona coluna Ano
        col_uf = [c for c in df.columns if c.strip().upper() == 'UF']
        if col_uf:
            df = df[df[col_uf[0]].str.strip() == filtro_uf].copy()
        df.insert(0, 'Ano', str(ano))

        grupos[(escopo, tabela)].append(df)

    caminhos = []
    nomes = {
        ('ESTDF', 'Receitas'): f'base_ESTDF_Receitas_{filtro_uf}.csv',
        ('ESTDF', 'Despesas'): f'base_ESTDF_Despesas_{filtro_uf}.csv',
        ('CAP',   'Receitas'): f'base_CAP_Receitas_{filtro_uf}.csv',
        ('CAP',   'Despesas'): f'base_CAP_Despesas_{filtro_uf}.csv',
    }
    for chave, frames in grupos.items():
        if not frames:
            continue
        df_consolidado = pd.concat(frames, ignore_index=True)
        path = os.path.join(dest, nomes[chave])
        df_consolidado.to_csv(path, index=False, encoding='utf-8-sig', sep=';')
        caminhos.append(path)

    return caminhos


def exportar_csv(dem1_estado, dem1_capital, dem2_estado, dem2_capital):
    """
    Exporta os 4 demonstrativos como arquivos CSV (UTF-8 BOM, separador ';').

    Retorna
    -------
    list[str] : caminhos dos arquivos gerados.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    arquivos = {
        'dem1_estado_MS.csv': dem1_estado,
        'dem1_capital_CampoGrande.csv': dem1_capital,
        'dem2_estado_MS.csv': dem2_estado,
        'dem2_capital_CampoGrande.csv': dem2_capital,
    }
    caminhos = []
    for nome, df in arquivos.items():
        path = os.path.join(OUTPUT_DIR, nome)
        df.to_csv(path, encoding='utf-8-sig', sep=';', decimal=',')
        caminhos.append(path)
    return caminhos
