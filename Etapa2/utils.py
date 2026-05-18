"""Funções utilitárias de formatação e cálculo."""


def fmt_mi(v):
    """Formata valor em reais como string em R$ milhões (3 casas)."""
    if v is None:
        return "–"
    return f"{v/1e6:.3f}"


def pct(numerador, denominador, casas=2):
    """Retorna percentual arredondado ou None se denominador inválido."""
    if denominador and denominador != 0:
        return round(numerador / denominador * 100, casas)
    return None


def get(d, *keys, default=0.0):
    """Retorna o valor da primeira chave encontrada no dicionário."""
    for k in keys:
        if k in d:
            return d[k]
    return default


def rcl6(d):
    """Retorna RCL Ajustada (VI); se ausente, retorna RCL bruta."""
    if "rcl_ajustada" in d:
        return d["rcl_ajustada"]
    return get(d, "rcl")
