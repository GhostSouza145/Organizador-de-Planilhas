"""
modules/generica.py
────────────────────
Organizador genérico para qualquer tipo de planilha.
Aplica limpeza inteligente sem depender de nomes de colunas específicos.

Regras aplicadas:
  - Limpeza básica (vazios, duplicados, strip)
  - Auto-detecção e formatação de colunas de data
  - Auto-detecção e limpeza de colunas numéricas/monetárias
  - Padroniza texto (Title Case) em colunas de texto puro
  - Nomeia colunas sem nome ("Coluna 1", "Coluna 2"...)
  - Ordena pela primeira coluna de data encontrada (se houver)
"""

import pandas as pd
import re
from utils.helpers import (
    limpeza_basica, padronizar_data, limpar_valor_monetario,
    titulo_proprio, exportar_xlsx_estilizado,
)


# Padrões para auto-detecção de tipo de coluna
KEYWORDS_DATA = [
    "data", "date", "dt", "dia", "mes", "mês", "ano", "year",
    "month", "day", "vencimento", "prazo", "cadastro", "nascimento",
]
KEYWORDS_MOEDA = [
    "valor", "value", "preco", "preço", "price", "total", "amount",
    "custo", "cost", "salario", "salário", "salary", "desconto",
    "taxa", "vlr", "vl", "montante",
]
KEYWORDS_TEXTO_NOME = [
    "nome", "name", "descricao", "descrição", "description", "titulo",
    "título", "title", "produto", "product", "cliente", "client",
    "fornecedor", "categoria", "cargo", "cidade", "estado", "bairro",
]


def _detectar_tipo_coluna(nome_col: str, serie: pd.Series) -> str:
    """
    Heurística para detectar o tipo de conteúdo de uma coluna.
    Retorna: 'data', 'moeda', 'numero', 'texto_nome', 'texto', 'desconhecido'
    """
    nome_lower = str(nome_col).lower()

    # Por nome da coluna
    if any(k in nome_lower for k in KEYWORDS_DATA):
        return "data"
    if any(k in nome_lower for k in KEYWORDS_MOEDA):
        return "moeda"
    if any(k in nome_lower for k in KEYWORDS_TEXTO_NOME):
        return "texto_nome"

    # Por conteúdo da série (amostra de até 50 linhas não-nulas)
    sample = serie.dropna().astype(str).head(50)
    if sample.empty:
        return "desconhecido"

    # Tenta detectar datas
    data_count = 0
    for v in sample:
        if re.search(r"\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}", v):
            data_count += 1
    if data_count > len(sample) * 0.6:
        return "data"

    # Tenta detectar valores monetários
    moeda_count = 0
    for v in sample:
        v_clean = re.sub(r"[R$€£\s\.]", "", v).replace(",", ".")
        try:
            float(v_clean)
            moeda_count += 1
        except ValueError:
            pass
    if moeda_count > len(sample) * 0.7:
        return "moeda"

    # Verifica se é texto (palavras, não números)
    texto_count = sum(1 for v in sample if re.search(r"[a-zA-ZÀ-ú]", v))
    if texto_count > len(sample) * 0.5:
        return "texto"

    return "desconhecido"


def organizar_generica(caminho_entrada: str, caminho_saida: str) -> dict:
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        # Tenta detectar separador automaticamente
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    # ── Limpeza básica ────────────────────────────────────
    df, stats = limpeza_basica(df)

    # ── Renomeia colunas sem nome ─────────────────────────
    novas_colunas = []
    contador = 1
    for col in df.columns:
        nome_str = str(col).strip()
        if nome_str.startswith("Unnamed") or nome_str in ("", "nan", "None"):
            novas_colunas.append(f"Coluna {contador}")
            contador += 1
        else:
            novas_colunas.append(nome_str)
            contador += 1
    df.columns = novas_colunas

    # ── Processa cada coluna por tipo ─────────────────────
    col_data_principal = None  # Para ordenação final

    for col in df.columns:
        tipo = _detectar_tipo_coluna(col, df[col])

        if tipo == "data":
            df[col] = df[col].apply(padronizar_data)
            if col_data_principal is None:
                col_data_principal = col

        elif tipo == "moeda":
            # Verifica se tem símbolo monetário para formatar
            tem_moeda = df[col].astype(str).str.contains(r"[R$€£,]", regex=True).any()
            if tem_moeda:
                num = df[col].apply(limpar_valor_monetario)
                df[col] = num.apply(
                    lambda v: f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")
                    if pd.notna(v) else ""
                )
            else:
                # Apenas converte para numérico e arredonda
                df[col] = pd.to_numeric(
                    df[col].apply(lambda x: str(x).replace(",", ".") if pd.notna(x) else None),
                    errors="coerce"
                ).round(2)

        elif tipo in ("texto_nome", "texto"):
            # Aplica Title Case apenas se parecer ser nome/descrição
            if tipo == "texto_nome":
                df[col] = df[col].apply(titulo_proprio)
            # Demais textos: apenas limpa espaços extras
            else:
                df[col] = df[col].apply(
                    lambda x: str(x).strip() if pd.notna(x) and str(x).strip() not in ("nan","None","") else ""
                )

    # ── Ordena pela primeira coluna de data ───────────────
    if col_data_principal:
        try:
            df["__sort"] = pd.to_datetime(
                df[col_data_principal], format="%d/%m/%Y", errors="coerce"
            )
            df = df.sort_values("__sort", na_position="last")
            df = df.drop(columns=["__sort"])
        except Exception:
            pass
    else:
        # Ordena pela primeira coluna de texto se não tiver data
        text_cols = [c for c in df.columns if df[c].dtype == object]
        if text_cols:
            try:
                df = df.sort_values(text_cols[0], na_position="last")
            except Exception:
                pass

    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Planilha")
    return stats
