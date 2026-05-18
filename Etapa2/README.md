# Mini Ministerio da Fazenda - Etapa 2

## Diagnostico de Endividamento

Esta etapa analisa indicadores de endividamento, liquidez e composicao da divida para Mato Grosso do Sul e Campo Grande, usando bases do Siconfi/Finbra e demonstrativos de Divida Consolidada Liquida.

## Estrutura

```text
Etapa2/
  base_dados/
    Divida Consolidada Liquida - Capitais - 2019.xlsx
    Divida Consolidada Liquida - Estados.DF - 2019.xlsx
    finbraRGF-2019.csv
    ...
  Material_entrega/
    1. base_dados_csv/
    2. graficos_e_tabelas/
    marcos_souza_quadro_estatistico_E2_E4.xlsx
    marcos_souza_relatorio_E2_E4.pdf
  Output/
  config.py
  extrair.py
  exportar.py
  processar_divida_etapa2.py
  generate_relatorio_endividamento_pdf.py
  utils.py
```

## Scripts Principais

- `extrair.py`: extrai e organiza dados das planilhas e CSVs de origem.
- `processar_divida_etapa2.py`: calcula indicadores e consolida a base da etapa.
- `exportar.py`: exporta tabelas e quadros estatisticos.
- `generate_relatorio_endividamento_pdf.py`: gera o relatorio final em PDF.

## Entradas

- Planilhas de Divida Consolidada Liquida para capitais e estados/DF, 2019 a 2024.
- Arquivos `finbraRGF-2019.csv` a `finbraRGF-2024.csv`.

## Saidas

- Bases CSV consolidadas em `Material_entrega/1. base_dados_csv/`.
- Graficos e tabelas em `Material_entrega/2. graficos_e_tabelas/`.
- Relatorio final em PDF e quadro estatistico em XLSX.

## Como Reproduzir

Execute os scripts a partir da pasta `Etapa2`:

```bash
python processar_divida_etapa2.py
python exportar.py
python generate_relatorio_endividamento_pdf.py
```

## Observacao

Os arquivos em `Doc_orientacoes/` ficam fora do repositorio por padrao, pois sao materiais de apoio da disciplina.

