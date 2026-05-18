# Mini Ministerio da Fazenda - Etapa 3

## Analise Econometrica

Esta etapa monta uma base municipio-ano para Mato Grosso do Sul e estima a relacao entre politica fiscal municipal e crescimento economico. O foco esta em preparar as variaveis do modelo, gerar estatisticas descritivas e executar regressao OLS/MQO.

## Estrutura

```text
Etapa3/
  base_dados/
    Despesas/
    Receitas/
    PIB/
    populacao/
    distancias-municipios-capital/
  Material_entrega/
    outputs/
    scripts/
      01_preparar_base.py
      02_analise_descritiva.ipynb
      03_regressao_ols.ipynb
      pipeline/
```

## Pipeline

O script principal e:

```bash
python Etapa3/Material_entrega/scripts/01_preparar_base.py
```

Ele consolida:

- receitas municipais Siconfi/Finbra;
- despesas municipais Siconfi/Finbra;
- PIB municipal do IBGE;
- populacao municipal DATASUS/TABNET;
- area territorial municipal;
- distancia ate Campo Grande pela base REGIC 2018.

## Notebooks

- `02_analise_descritiva.ipynb`: estatisticas descritivas, matriz de correlacao, outliers e visualizacoes com Plotly.
- `03_regressao_ols.ipynb`: estimacao OLS/MQO e diagnosticos econometricos.

## Principais Saidas

Arquivos em `Material_entrega/outputs/`:

- `base_etapa3_ms.csv`
- `base_distancia_capital_ms.csv`
- `estatisticas_descritivas.csv`
- `matriz_correlacao.csv`
- `outliers_potenciais.csv`
- `analise_descritiva_resultados.xlsx`
- `regressao_ols_resultados.xlsx`

## Variaveis Centrais

- `crescimento_pib`
- `log_pib_per_capita`
- `investimento_pib`
- `gasto_pessoal_pib`
- `receita_corrente_pib`
- `crescimento_populacao`
- `area_por_habitante`
- `distancia_capital_km`
- `dummy_pequeno_municipio`
- `interacao_pequeno_pessoal`

## Como Reproduzir

Na raiz do projeto:

```bash
python Etapa3/Material_entrega/scripts/01_preparar_base.py
```

Depois, abra os notebooks com o kernel `Python (work)` ou com um ambiente que contenha as dependencias listadas em `requirements.txt`.

## Observacao

Os arquivos em `Doc_orientacoes/` ficam fora do repositorio por padrao, pois sao materiais de apoio da disciplina.

