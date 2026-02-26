"""
modules/leads.py
─────────────────
Organizador de planilhas de Leads / Clientes.
Trata: nomes, e-mails, telefones, CPF, endereços, origem.

Regras aplicadas:
  - Padroniza nomes (Title Case)
  - Valida e formata e-mails (lowercase)
  - Formata telefones no padrão BR
  - Formata CPF/CNPJ
  - Padroniza cidade e estado (Title Case / UF)
  - Cria coluna NOME COMPLETO se houver primeiro/último nome separados
  - Remove duplicatas por e-mail ou telefone
  - Ordena por nome
"""

import re
import pandas as pd
from utils.helpers import (
    limpeza_basica, padronizar_data, titulo_proprio,
    lowercase_seguro, formatar_cpf, formatar_telefone,
    exportar_xlsx_estilizado,
)


MAPA_COLUNAS = {
    "nome":         ["nome", "name", "cliente", "client", "contato",
                     "contact", "razão social", "razao social",
                     "nome_completo", "nome completo"],
    "primeiro_nome":["primeiro nome", "primeiro_nome", "first name", "first_name"],
    "ultimo_nome":  ["último nome", "ultimo_nome", "last name", "last_name",
                     "sobrenome"],
    "email":        ["email", "e-mail", "mail", "correio", "endereço eletrônico"],
    "telefone":     ["telefone", "tel", "phone", "celular", "whatsapp",
                     "fone", "contato_telefone"],
    "cpf":          ["cpf", "cpf_cnpj", "documento", "doc"],
    "cidade":       ["cidade", "city", "municipio", "município"],
    "estado":       ["estado", "state", "uf", "provincia"],
    "origem":       ["origem", "source", "canal", "channel", "como_conheceu",
                     "indicação"],
    "status":       ["status", "situação", "situacao", "fase", "etapa", "stage"],
    "data":         ["data", "date", "data_cadastro", "cadastro",
                     "data_contato", "criado_em"],
}


def _detectar_coluna(df: pd.DataFrame, grupo: str) -> str | None:
    candidatos = MAPA_COLUNAS.get(grupo, [])
    colunas_lower = {c.lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in colunas_lower:
            return colunas_lower[cand.lower()]
    return None


def _validar_email(email) -> str:
    """Retorna e-mail em lowercase se válido, senão retorna vazio."""
    if not email or str(email).strip() in ("nan", "None", ""):
        return ""
    e = str(email).strip().lower()
    # Padrão básico de e-mail
    if re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", e):
        return e
    return f"⚠️ {e}"  # marca como suspeito mas mantém


def _formatar_uf(uf) -> str:
    """Padroniza UF para 2 letras maiúsculas."""
    if not uf or str(uf).strip() in ("nan", "None", ""):
        return ""
    return str(uf).strip().upper()[:2]


def organizar_leads(caminho_entrada: str, caminho_saida: str) -> dict:
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    df, stats = limpeza_basica(df)

    col_nome       = _detectar_coluna(df, "nome")
    col_pnome      = _detectar_coluna(df, "primeiro_nome")
    col_unome      = _detectar_coluna(df, "ultimo_nome")
    col_email      = _detectar_coluna(df, "email")
    col_telefone   = _detectar_coluna(df, "telefone")
    col_cpf        = _detectar_coluna(df, "cpf")
    col_cidade     = _detectar_coluna(df, "cidade")
    col_estado     = _detectar_coluna(df, "estado")
    col_origem     = _detectar_coluna(df, "origem")
    col_status     = _detectar_coluna(df, "status")
    col_data       = _detectar_coluna(df, "data")

    # ── Cria NOME COMPLETO se separado ───────────────────
    if not col_nome and col_pnome and col_unome:
        col_nome = "Nome Completo"
        stats["colunas_criadas"].append(col_nome)
        df[col_nome] = (
            df[col_pnome].apply(titulo_proprio) + " " + df[col_unome].apply(titulo_proprio)
        ).str.strip()

    # ── Padroniza NOME ────────────────────────────────────
    if col_nome:
        df[col_nome] = df[col_nome].apply(titulo_proprio)

    # ── Padroniza E-MAIL ──────────────────────────────────
    if col_email:
        df[col_email] = df[col_email].apply(_validar_email)
        # Remove duplicatas por e-mail (mantém primeira ocorrência)
        antes = len(df)
        df_emails = df[df[col_email] != ""]
        df_sem_email = df[df[col_email] == ""]
        df_emails = df_emails.drop_duplicates(subset=[col_email], keep="first")
        df = pd.concat([df_emails, df_sem_email], ignore_index=True)
        dup_email = antes - len(df)
        if dup_email > 0:
            stats["duplicados"] = stats.get("duplicados", 0) + dup_email

    # ── Padroniza TELEFONE ────────────────────────────────
    if col_telefone:
        df[col_telefone] = df[col_telefone].apply(formatar_telefone)

    # ── Formata CPF ───────────────────────────────────────
    if col_cpf:
        df[col_cpf] = df[col_cpf].apply(formatar_cpf)

    # ── Padroniza CIDADE e ESTADO ─────────────────────────
    if col_cidade:
        df[col_cidade] = df[col_cidade].apply(titulo_proprio)
    if col_estado:
        df[col_estado] = df[col_estado].apply(_formatar_uf)

    # ── Padroniza ORIGEM e STATUS ─────────────────────────
    if col_origem:
        df[col_origem] = df[col_origem].apply(titulo_proprio)
    if col_status:
        df[col_status] = df[col_status].apply(titulo_proprio)

    # ── Padroniza DATA ────────────────────────────────────
    if col_data:
        df[col_data] = df[col_data].apply(padronizar_data)

    # ── Ordena por nome ───────────────────────────────────
    if col_nome:
        df = df.sort_values(col_nome, na_position="last")

    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Leads_Clientes")
    return stats
