"""
modules/vendas.py
──────────────────
Organizador de planilhas de vendas.
Trata: pedidos, clientes, produtos, valores, datas.

Regras aplicadas:
  - Padroniza datas de pedido
  - Normaliza nomes de clientes (Title Case)
  - Limpa e formata valores monetários
  - Cria coluna de STATUS simplificado se ausente
  - Cria coluna de MÊS/ANO para análise
  - Ordena por data decrescente (mais recente primeiro)
  - Remove duplicatas e linhas vazias
"""

import pandas as pd
from utils.helpers import (
    limpeza_basica, padronizar_data, limpar_valor_monetario,
    titulo_proprio, uppercase_seguro, exportar_xlsx_estilizado,
)


MAPA_COLUNAS = {
    "data":         ["data", "date", "dt", "data_pedido", "data pedido",
                     "data_venda", "data venda", "data_emissão", "emissão"],
    "pedido":       ["pedido", "order", "numero_pedido", "nº pedido", "nro",
                     "id_pedido", "id pedido", "cod_pedido", "número"],
    "cliente":      ["cliente", "client", "customer", "comprador",
                     "nome_cliente", "nome cliente", "razão social"],
    "produto":      ["produto", "product", "item", "descrição", "descricao",
                     "mercadoria", "serviço"],
    "quantidade":   ["quantidade", "qtd", "qty", "quantity", "qtde", "unidades"],
    "valor":        ["valor", "value", "amount", "total", "subtotal",
                     "vlr_total", "valor total", "montante"],
    "status":       ["status", "situação", "situacao", "state",
                     "andamento", "fase"],
    "vendedor":     ["vendedor", "seller", "representative", "representante",
                     "consultor", "agente"],
}


def _detectar_coluna(df: pd.DataFrame, grupo: str) -> str | None:
    candidatos = MAPA_COLUNAS.get(grupo, [])
    colunas_lower = {c.lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in colunas_lower:
            return colunas_lower[cand.lower()]
    return None


# Normaliza status comuns para padrão legível
STATUS_MAP = {
    "pago": "✅ Pago", "paid": "✅ Pago", "aprovado": "✅ Pago",
    "pendente": "⏳ Pendente", "pending": "⏳ Pendente", "aguardando": "⏳ Pendente",
    "cancelado": "❌ Cancelado", "canceled": "❌ Cancelado", "cancelled": "❌ Cancelado",
    "enviado": "🚚 Enviado", "shipped": "🚚 Enviado", "em transito": "🚚 Enviado",
    "entregue": "📦 Entregue", "delivered": "📦 Entregue",
    "devolvido": "↩️ Devolvido", "returned": "↩️ Devolvido",
}


def _normalizar_status(valor) -> str:
    if not valor or str(valor).strip() in ("nan", "None", ""):
        return "❓ Não informado"
    lower = str(valor).strip().lower()
    return STATUS_MAP.get(lower, titulo_proprio(str(valor)))


def organizar_vendas(caminho_entrada: str, caminho_saida: str) -> dict:
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    df, stats = limpeza_basica(df)

    col_data       = _detectar_coluna(df, "data")
    col_pedido     = _detectar_coluna(df, "pedido")
    col_cliente    = _detectar_coluna(df, "cliente")
    col_produto    = _detectar_coluna(df, "produto")
    col_quantidade = _detectar_coluna(df, "quantidade")
    col_valor      = _detectar_coluna(df, "valor")
    col_status     = _detectar_coluna(df, "status")
    col_vendedor   = _detectar_coluna(df, "vendedor")

    # ── Padroniza DATA ───────────────────────────────────
    if col_data:
        df[col_data] = df[col_data].apply(padronizar_data)

    # ── Padroniza CLIENTE ────────────────────────────────
    if col_cliente:
        df[col_cliente] = df[col_cliente].apply(titulo_proprio)

    # ── Padroniza PRODUTO ────────────────────────────────
    if col_produto:
        df[col_produto] = df[col_produto].apply(titulo_proprio)

    # ── Padroniza VENDEDOR ───────────────────────────────
    if col_vendedor:
        df[col_vendedor] = df[col_vendedor].apply(titulo_proprio)

    # ── Limpa e formata VALOR ────────────────────────────
    if col_valor:
        df[col_valor] = df[col_valor].apply(limpar_valor_monetario)
        df[col_valor] = df[col_valor].apply(
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notna(v) else ""
        )

    # ── Normaliza STATUS ─────────────────────────────────
    if col_status:
        df[col_status] = df[col_status].apply(_normalizar_status)

    # ── Converte QUANTIDADE ──────────────────────────────
    if col_quantidade:
        df[col_quantidade] = pd.to_numeric(
            df[col_quantidade].apply(lambda x: str(x).replace(",", ".") if pd.notna(x) else None),
            errors="coerce"
        )

    # ── Cria coluna MÊS/ANO ──────────────────────────────
    if col_data:
        try:
            col_mes = "Mês/Ano"
            stats["colunas_criadas"].append(col_mes)
            df[col_mes] = pd.to_datetime(
                df[col_data], format="%d/%m/%Y", errors="coerce"
            ).dt.strftime("%m/%Y")
        except Exception:
            pass

    # ── Ordena por data (mais recente primeiro) ──────────
    if col_data:
        try:
            df["__sort_dt"] = pd.to_datetime(df[col_data], format="%d/%m/%Y", errors="coerce")
            df = df.sort_values("__sort_dt", ascending=False, na_position="last")
            df = df.drop(columns=["__sort_dt"])
        except Exception:
            pass

    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Vendas")
    return stats
