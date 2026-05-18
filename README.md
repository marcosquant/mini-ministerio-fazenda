# Mini Ministerio da Fazenda

Projeto academico do Eixo 4 da PUC Minas para analise fiscal, endividamento e modelagem econometrica com dados publicos de financas municipais e estaduais.

## Estrutura

- `Etapa1/`: diagnostico fiscal de Mato Grosso do Sul e Campo Grande.
- `Etapa2/`: diagnostico de endividamento e liquidez.
- `Etapa3/`: preparacao da base econometrica, analise descritiva e regressao.

## Fontes de dados

O projeto usa bases publicas do Siconfi/Finbra, IBGE, DATASUS/TABNET e REGIC/IBGE, alem de planilhas derivadas geradas pelos scripts de cada etapa.

## Como reproduzir

1. Crie o ambiente Python:

```bash
conda env create -f environment.yml
conda activate work
```

2. Execute a preparacao principal da Etapa 3:

```bash
python Etapa3/Material_entrega/scripts/01_preparar_base.py
```

3. Abra os notebooks:

```bash
jupyter notebook
```

Notebooks principais:

- `Etapa3/Material_entrega/scripts/02_analise_descritiva.ipynb`
- `Etapa3/Material_entrega/scripts/03_regressao_ols.ipynb`

## Observacoes

- O repositorio foi preparado para ser privado.
- Os arquivos de orientacao da disciplina em `Doc_orientacoes/` ficam fora do Git por padrao.
- Arquivos grandes como PDFs, DOCX, XLSX, imagens e ZIPs sao rastreados com Git LFS.

