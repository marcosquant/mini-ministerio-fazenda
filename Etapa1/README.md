# Mini Ministério da Fazenda — Etapa 1

**Análise Fiscal: Mato Grosso do Sul e Campo Grande (2019–2024)**

Projeto acadêmico da PUC Minas (4º Semestre — Eixo 4).
Processa dados FINBRA/SICONFI para gerar demonstrativos fiscais do **Estado do Mato Grosso do Sul** e de sua capital **Campo Grande**, no período 2019–2024.

---

## Estrutura do Projeto

```
Etapa 1/
├── main.py                    # Ponto de entrada
├── src/
│   ├── __init__.py
│   ├── config.py              # Configurações globais (diretórios, anos, UF, cores)
│   ├── carregador.py          # Leitura dos ZIPs FINBRA e extração de valores
│   ├── calculadora.py         # Cálculos fiscais por ano (estado e capital)
│   ├── demonstrativos.py      # Construção dos DataFrames dos demonstrativos
│   ├── exportador.py          # Exportação para Excel (xlsxwriter) e CSV
│   └── graficos.py            # Geração dos 8 gráficos (matplotlib)
├── base_dados/                # Arquivos ZIP do FINBRA (não versionados)
│   ├── finbra_ESTDF_ReceitasOrcamentarias(AnexoI-C)_2019.zip
│   ├── finbra_ESTDF_DespesasOrcamentarias(AnexoI-D)_2019.zip
│   ├── finbra_CAP_ReceitasOrcamentarias(AnexoI-C)_2019.zip
│   ├── finbra_CAP_DespesasOrcamentarias(AnexoI-D)_2019.zip
│   └── ... (mesmo padrão para 2020–2024)
└── output/                    # Arquivos gerados (criado automaticamente)
    ├── demonstrativos_fiscais_MS_2019_2024.xlsx
    ├── dem1_estado_MS.csv
    ├── dem1_capital_CampoGrande.csv
    ├── dem2_estado_MS.csv
    ├── dem2_capital_CampoGrande.csv
    ├── grafico1_resultado_primario_estado.png
    ├── grafico2_resultado_primario_capital.png
    ├── grafico3_rec_desp_primarias_estado.png
    ├── grafico4_rec_desp_primarias_capital.png
    ├── grafico5_composicao_despesas_estado.png
    ├── grafico6_composicao_receitas_estado.png
    ├── grafico7_resultado_orcamentario.png
    └── grafico8_evolucao_receitas_despesas_estado.png
```

---

## Como Executar

### Pré-requisitos

- Python 3.8+
- Pacotes: `pandas`, `numpy`, `matplotlib`, `xlsxwriter`

```bash
pip install pandas numpy matplotlib xlsxwriter
```

### Execução

```bash
python main.py
```

O script processa os dados, exibe progresso no terminal e salva todos os arquivos na pasta `output/`.

---

## Fonte dos Dados

**FINBRA (Finanças do Brasil)** — Sistema SICONFI do Tesouro Nacional
Portal: [siconfi.tesouro.gov.br](https://siconfi.tesouro.gov.br)

Arquivos utilizados (por ano, 2019–2024):
- `finbra_ESTDF_ReceitasOrcamentarias(AnexoI-C)_{ano}.zip` — Receitas dos Estados
- `finbra_ESTDF_DespesasOrcamentarias(AnexoI-D)_{ano}.zip` — Despesas dos Estados
- `finbra_CAP_ReceitasOrcamentarias(AnexoI-C)_{ano}.zip` — Receitas das Capitais
- `finbra_CAP_DespesasOrcamentarias(AnexoI-D)_{ano}.zip` — Despesas das Capitais

Cada ZIP contém `finbra.csv` (encoding `latin-1`, separador `;`, 3 linhas de cabeçalho de metadados antes da linha de colunas).

---

## Metodologia

### Receitas Líquidas

Adota-se a metodologia padrão FINBRA/SICONFI:

```
Receita Líquida = Receitas Brutas Realizadas
                − Deduções FUNDEB
                − Deduções Transferências Constitucionais
```

> A coluna "Outras Deduções da Receita" (introduzida em 2022, relativa à isenção do ICMS-combustíveis pela Lei 194/2022) é **excluída** para manter a série histórica consistente.

### Despesas

Utiliza-se a coluna **Despesas Empenhadas** (regime de competência orçamentária).

### Mudança de Códigos em 2022

O FINBRA reestruturou os códigos de contas a partir de 2022. O sistema lida com isso via lista de fallback nos padrões de conta:

| Conta | Código até 2021 | Código 2022+ |
|---|---|---|
| IPVA (Estado) | `1.1.1.8.01.2.0` | `1.1.1.2.51.0.0` |
| ICMS (Estado) | `1.1.1.8.02.{1,2}.0` | `1.1.1.4.50.{1,2}.0` |
| FPE (Estado) | `1.7.1.8.01.1.0` | `1.7.1.1.50.0.0` |
| IPTU (Capital) | `1.1.1.8.01.1.0` | `1.1.1.2.50.0.0` |
| ISS (Capital) | `1.1.1.8.02.3.0` | `1.1.1.4.51.1.0` |
| FPM (Capital) | `1.7.1.8.01.2.0` | `1.7.1.1.51.0.0` |

### Demonstrativos

#### 1º Demonstrativo — Resultado Primário

| Saldo | Descrição | Fórmula |
|---|---|---|
| A | Receitas Primárias Correntes | Rec. Correntes − Rec. Financeiras Correntes |
| B | Receitas Primárias de Capital | Rec. Capital − (Op. Crédito + Alienação + Amort. recebidas) |
| C | **Receita Primária Total** | A + B |
| D | Despesas Primárias Correntes | Desp. Correntes − Juros e Encargos |
| E | Despesas Primárias de Capital | Desp. Capital − Desp. Financeiras de Capital |
| F | **Despesa Primária Total** | D + E |
| **G** | **Resultado Primário** | **C − F** |

#### 2º Demonstrativo — Resultado Orçamentário

```
Resultado Orçamentário = Resultado Primário
                       + Receitas Financeiras
                       − Despesas Financeiras
```

Onde:
- **Receitas Financeiras** = Valores Mobiliários + Operações de Crédito + Alienação de Bens + Amortizações recebidas
- **Despesas Financeiras** = Juros e Encargos da Dívida + Amortização da Dívida

---

## Saídas Geradas

### Excel — `demonstrativos_fiscais_MS_2019_2024.xlsx`

4 abas com formatação condicional:
- `Dem1_Estado_MS` — 1º Demonstrativo do Estado do MS
- `Dem1_Capital_CG` — 1º Demonstrativo de Campo Grande
- `Dem2_Estado_MS` — 2º Demonstrativo do Estado do MS
- `Dem2_Capital_CG` — 2º Demonstrativo de Campo Grande

Codificação de cores:
- 🔵 Azul escuro: seções principais
- 🟢 Verde: saldos (A, B, D, E) e resultados positivos
- 🔴 Vermelho: resultados negativos

### CSVs

Valores em R$ milhões, separador `;`, decimal `,`, encoding UTF-8 BOM.

### Gráficos (PNG, 150 DPI)

| Arquivo | Descrição |
|---|---|
| `grafico1_resultado_primario_estado.png` | Barras: Resultado Primário — Estado MS |
| `grafico2_resultado_primario_capital.png` | Barras: Resultado Primário — Campo Grande |
| `grafico3_rec_desp_primarias_estado.png` | Barras agrupadas: Rec. vs Desp. — Estado MS |
| `grafico4_rec_desp_primarias_capital.png` | Barras agrupadas: Rec. vs Desp. — Campo Grande |
| `grafico5_composicao_despesas_estado.png` | Pizza: Composição Despesas Estado (2019 e 2024) |
| `grafico6_composicao_receitas_estado.png` | Pizza: Composição Receitas Estado (2019 e 2024) |
| `grafico7_resultado_orcamentario.png` | Barras duplas: Resultado Orçamentário |
| `grafico8_evolucao_receitas_despesas_estado.png` | Linhas: ICMS, Pessoal e Rec. Total — Estado MS |

---

## Módulos

### `src/config.py`
Constantes globais: caminhos de diretório, lista de anos, UF e paleta de cores.

### `src/carregador.py`
- `ler_finbra(tipo_escopo, tipo_tabela, ano, filtro_uf)` — lê ZIP, decodifica latin-1, retorna DataFrame filtrado.
- `get_valor(df, coluna_filtro, conta_pattern)` — extrai valor de uma conta. `conta_pattern` pode ser lista para fallback de códigos.
- `get_valor_liquido(df, conta_pattern)` — retorna receita líquida (Brutas − FUNDEB − Transf. Constitucionais).

### `src/calculadora.py`
- `calcular_estado(ano)` — processa todos os indicadores do Estado para um ano.
- `calcular_capital(ano)` — processa todos os indicadores da Capital para um ano.

Ambas retornam `dict` com ~30 chaves de métricas fiscais.

### `src/demonstrativos.py`
- `build_demonstrativo1_estado(resultados)` — constrói DataFrame do 1º Demonstrativo (Estado).
- `build_demonstrativo1_capital(resultados)` — constrói DataFrame do 1º Demonstrativo (Capital).
- `build_demonstrativo2(resultados)` — constrói DataFrame do 2º Demonstrativo (Estado ou Capital).

### `src/exportador.py`
- `exportar_excel(dem1_est, dem1_cap, dem2_est, dem2_cap)` — gera Excel formatado com xlsxwriter.
- `exportar_csv(dem1_est, dem1_cap, dem2_est, dem2_cap)` — gera 4 CSVs.

### `src/graficos.py`
- `gerar_todos(resultados_estado, resultados_capital)` — gera os 8 gráficos e retorna lista de caminhos.
- Funções individuais: `grafico1_*` a `grafico8_*` para geração granular.

---

## Observações Técnicas

- Os déficits primários de Campo Grande (−R$66 mi a −R$736 mi) são consistentes com a estrutura fiscal do município, que utiliza operações de crédito para financiar investimentos.
- A Receita de Capital do Estado MS inclui volumes expressivos de Operações de Crédito em determinados anos, ampliando o Resultado Orçamentário acima do Primário.
- Todos os valores estão em **R$ milhões** nos arquivos de saída.
