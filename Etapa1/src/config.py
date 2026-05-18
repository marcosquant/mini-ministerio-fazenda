"""
Configurações globais do projeto.
"""

import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'base_dados')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

ANOS = [2019, 2020, 2021, 2022, 2023, 2024]
UF = 'MS'

CORES = {
    'azul': '#1F4E79',
    'verde': '#375623',
    'laranja': '#C55A11',
    'vermelho': '#C00000',
    'cinza': '#595959',
    'azul_claro': '#2E75B6',
    'verde_claro': '#70AD47',
}
