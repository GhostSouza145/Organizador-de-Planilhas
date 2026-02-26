"""
modules/estoque.py
───────────────────
Organizador de planilhas de estoque.
Trata: produtos, SKU, quantidades, preços, fornecedores.

Regras aplicadas:
  - Padroniza nomes de produtos (Title Case)
  - Normaliza códigos/SKU (MAIÚSCULAS, sem espaços extras)
  - Trata valores numéricos (quantidade, preço)
  - Cria coluna de "Status Estoque" (Crítico / Baixo / Normal / Excesso)
  - Cria coluna de Valor Total (Qtd × Preço) se possível
  - Ordena por nome do produto
  - Remove duplicatas e linhas vazias
"""

import pandas as pd
from utils.helpers import (
    limpeza_basica, limpar_valor_monetario,
    titulo_proprio, uppercase_seguro, exportar_xlsx_estilizado,
)


MAPA_COLUNAS = {
    "produto":      ["produto", "product", "item", "nome", "name",
                     "descrição", "descricao", "description", "mercadoria"],
    "codigo":       ["codigo", "código", "code", "sku", "ref", "referência",
                     "referencia", "cod", "id_produto", "id produto"],
    "quantidade":   ["quantidade", "qtd", "qty", "quantity", "estoque",
                     "saldo", "disponivel", "disponível", "qtde"],
    "preco":        ["preco", "preço", "price", "valor", "custo", "cost",
                     "preco_unitario", "preço unitário", "vlr_unit"],
    "minimo":       ["minimo", "mínimo", "min", "estoque_min", "qtd_min",
                     "ponto_pedido", "reposição"],
    "fornecedor":   ["fornecedor", "supplier", "vendor", "fabricante",
                     "marca", "brand"],
    "categoria":    ["categoria", "category", "grupo", "group",
                     "departamento", "setor"],
}


def _detectar_coluna(df: pd.DataFrame, grupo: str) -> str | None:
    candidatos = MAPA_COLUNAS.get(grupo, [])
    colunas_lower = {c.lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in colunas_lower:
            return colunas_lower[cand.lower()]
    return None


def _status_estoque(qtd, minimo=None) -> str:
    """Classifica o status do estoque."""
    try:
        qtd = float(qtd)
    except (TypeError, ValueError):
        return "Não informado"

    if minimo is not None:
        try:
            min_val = float(minimo)
            if qtd <= 0:
                return "🔴 Sem Estoque"
            if qtd <= min_val * 0.5:
                return "🟠 Crítico"
            if qtd <= min_val:
                return "🟡 Baixo"
            if qtd > min_val * 5:
                return "🔵 Excesso"
            return "🟢 Normal"
        except Exception:
            pass

    # Sem mínimo definido: usa heurística absoluta
    if qtd <= 0:
        return "🔴 Sem Estoque"
    if qtd <= 5:
        return "🟠 Crítico"
    if qtd <= 20:
        return "🟡 Baixo"
    return "🟢 Normal"


def organizar_estoque(caminho_entrada: str, caminho_saida: str) -> dict:
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    df, stats = limpeza_basica(df)

    # ── Detecta colunas ──────────────────────────────────
    col_produto    = _detectar_coluna(df, "produto")
    col_codigo     = _detectar_coluna(df, "codigo")
    col_quantidade = _detectar_coluna(df, "quantidade")
    col_preco      = _detectar_coluna(df, "preco")
    col_minimo     = _detectar_coluna(df, "minimo")
    col_fornecedor = _detectar_coluna(df, "fornecedor")
    col_categoria  = _detectar_coluna(df, "categoria")

    # ── Padroniza PRODUTO ────────────────────────────────
    if col_produto:
        df[col_produto] = df[col_produto].apply(titulo_proprio)

    # ── Padroniza CÓDIGO / SKU ───────────────────────────
    if col_codigo:
        df[col_codigo] = df[col_codigo].apply(
            lambda x: str(x).strip().upper().replace(" ", "") if pd.notna(x) else ""
        )

    # ── Converte QUANTIDADE para numérico ────────────────
    if col_quantidade:
        df[col_quantidade] = pd.to_numeric(
            df[col_quantidade].apply(lambda x: str(x).replace(",", ".") if pd.notna(x) else None),
            errors="coerce"
        )

    # ── Converte PREÇO para numérico e formata ───────────
    if col_preco:
        df[col_preco] = df[col_preco].apply(limpar_valor_monetario)

    # ── Cria coluna VALOR TOTAL ──────────────────────────
    if col_quantidade and col_preco:
        col_vt = "Valor Total (R$)"
        stats["colunas_criadas"].append(col_vt)
        df[col_vt] = (df[col_quantidade].fillna(0) * df[col_preco].fillna(0)).round(2)
        # Formata
        df[col_vt] = df[col_vt].apply(
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    # ── Cria coluna STATUS ESTOQUE ───────────────────────
    if col_quantidade:
        col_status = "Status Estoque"
        stats["colunas_criadas"].append(col_status)
        df[col_status] = df.apply(
            lambda row: _status_estoque(
                row[col_quantidade],
                row[col_minimo] if col_minimo else None
            ),
            axis=1
        )

    # ── Formata preço como moeda ─────────────────────────
    if col_preco:
        df[col_preco] = df[col_preco].apply(
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notna(v) else ""
        )

    # ── Padroniza FORNECEDOR e CATEGORIA ─────────────────
    if col_fornecedor:
        df[col_fornecedor] = df[col_fornecedor].apply(titulo_proprio)
    if col_categoria:
        df[col_categoria] = df[col_categoria].apply(titulo_proprio)

    # ── Ordena por produto ───────────────────────────────
    if col_produto:
        df = df.sort_values(col_produto, na_position="last")

    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Estoque")
    return stats
