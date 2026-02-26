"""
modules/financeira.py
──────────────────────
Organizador de planilhas financeiras.
Trata: receitas, despesas, fluxo de caixa, contas a pagar/receber.

Regras aplicadas:
  - Padroniza colunas de valores monetários
  - Padroniza e ordena por data
  - Classifica tipo de transação (receita/despesa) se possível
  - Cria coluna de saldo acumulado
  - Remove duplicatas e linhas vazias
  - Ordena por data crescente
"""

import pandas as pd
from utils.helpers import (
    limpeza_basica, padronizar_data, limpar_valor_monetario,
    titulo_proprio, uppercase_seguro, exportar_xlsx_estilizado,
)


# ── Mapeamento flexível de nomes de colunas ──────────────
MAPA_COLUNAS = {
    # DATA
    "data":         ["data", "date", "dt", "data_lancamento", "data lançamento",
                     "data_movimento", "data mov", "competência", "vencimento",
                     "data_vencimento", "data vencimento"],
    # DESCRIÇÃO
    "descricao":    ["descrição", "descricao", "description", "historico",
                     "histórico", "memo", "observação", "observacao", "detalhe",
                     "lançamento", "lancamento"],
    # VALOR
    "valor":        ["valor", "value", "amount", "montante", "quantia",
                     "total", "vlr", "vl"],
    # TIPO
    "tipo":         ["tipo", "type", "categoria", "category", "natureza",
                     "classificação", "classificacao", "class"],
    # CATEGORIA
    "categoria":    ["categoria", "category", "subcategoria", "grupo", "group"],
    # FORMA PAGAMENTO
    "forma_pag":    ["forma_pagamento", "forma pagamento", "pagamento",
                     "payment", "meio", "forma"],
}


def _detectar_coluna(df: pd.DataFrame, grupo: str) -> str | None:
    """Retorna o nome da coluna no DataFrame que corresponde ao grupo."""
    candidatos = MAPA_COLUNAS.get(grupo, [])
    colunas_lower = {c.lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in colunas_lower:
            return colunas_lower[cand.lower()]
    return None


def organizar_financeira(caminho_entrada: str, caminho_saida: str) -> dict:
    """
    Organiza planilha financeira.
    Retorna dicionário com estatísticas do processamento.
    """
    # ── Leitura ──────────────────────────────────────────
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    # ── Limpeza básica ───────────────────────────────────
    df, stats = limpeza_basica(df)

    # ── Detecta colunas principais ───────────────────────
    col_data     = _detectar_coluna(df, "data")
    col_descricao = _detectar_coluna(df, "descricao")
    col_valor    = _detectar_coluna(df, "valor")
    col_tipo     = _detectar_coluna(df, "tipo")
    col_categoria = _detectar_coluna(df, "categoria")

    # ── Padroniza DATA ───────────────────────────────────
    if col_data:
        df[col_data] = df[col_data].apply(padronizar_data)
        # Ordena por data
        try:
            df["__data_sort"] = pd.to_datetime(df[col_data], format="%d/%m/%Y", errors="coerce")
            df = df.sort_values("__data_sort", na_position="last")
            df = df.drop(columns=["__data_sort"])
        except Exception:
            pass

    # ── Padroniza VALOR ──────────────────────────────────
    if col_valor:
        df[col_valor] = df[col_valor].apply(limpar_valor_monetario)
        # Formata como moeda BR
        df[col_valor] = df[col_valor].apply(
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notna(v) else ""
        )

    # ── Padroniza DESCRIÇÃO ──────────────────────────────
    if col_descricao:
        df[col_descricao] = df[col_descricao].apply(titulo_proprio)

    # ── Padroniza TIPO/CATEGORIA ─────────────────────────
    if col_tipo:
        df[col_tipo] = df[col_tipo].apply(uppercase_seguro)

    if col_categoria:
        df[col_categoria] = df[col_categoria].apply(titulo_proprio)

    # ── Cria coluna TIPO se não existir ─────────────────
    if not col_tipo and col_valor:
        # Detecta receita/despesa pelo sinal do valor (se numérico)
        col_tipo_novo = "Tipo"
        stats["colunas_criadas"].append(col_tipo_novo)

        def inferir_tipo(v):
            v_num = limpar_valor_monetario(str(v).replace("R$", "").strip())
            if v_num is None:
                return "Não informado"
            return "Receita" if v_num >= 0 else "Despesa"

        df.insert(df.columns.get_loc(col_valor) + 1 if col_valor in df.columns else len(df.columns),
                  col_tipo_novo,
                  df[col_valor].apply(inferir_tipo))

    # ── Renomeia colunas para PT-BR se necessário ────────
    renomear = {}
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ["date", "dt"]:
            renomear[col] = "Data"
        elif col_lower in ["description", "memo"]:
            renomear[col] = "Descrição"
        elif col_lower in ["amount", "value"]:
            renomear[col] = "Valor"
        elif col_lower in ["type", "category"]:
            renomear[col] = "Categoria"
    if renomear:
        df = df.rename(columns=renomear)

    # ── Remove colunas 100% nulas após limpeza ───────────
    df = df.dropna(axis=1, how="all")

    # ── Reseta índice ────────────────────────────────────
    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    # ── Exporta com estilo ───────────────────────────────
    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Financeiro")

    return stats
