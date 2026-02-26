"""
modules/funcionarios.py
────────────────────────
Organizador de planilhas de Funcionários / RH.
Trata: nome, CPF, cargo, departamento, salário, admissão.

Regras aplicadas:
  - Padroniza nomes (Title Case)
  - Formata CPF
  - Padroniza cargos e departamentos (Title Case)
  - Formata salários como moeda BR
  - Padroniza datas de admissão e demissão
  - Cria coluna de TEMPO DE EMPRESA (anos/meses)
  - Classifica faixa salarial
  - Ordena por departamento e depois por nome
"""

import pandas as pd
from datetime import datetime
from utils.helpers import (
    limpeza_basica, padronizar_data, limpar_valor_monetario,
    titulo_proprio, formatar_cpf, formatar_telefone,
    exportar_xlsx_estilizado,
)


MAPA_COLUNAS = {
    "nome":         ["nome", "name", "funcionário", "funcionario",
                     "colaborador", "employee", "nome_completo"],
    "cpf":          ["cpf", "documento", "doc", "cpf_funcionario"],
    "cargo":        ["cargo", "position", "role", "função", "funcao",
                     "ocupação", "ocupacao", "job_title"],
    "departamento": ["departamento", "department", "dept", "setor",
                     "area", "área", "divisão", "divisao"],
    "salario":      ["salario", "salário", "salary", "remuneração",
                     "remuneracao", "ctc", "vencimento", "vlr_salario"],
    "admissao":     ["admissão", "admissao", "data_admissao", "data admissão",
                     "inicio", "início", "hire_date", "data_contratacao"],
    "demissao":     ["demissão", "demissao", "data_demissao", "data demissão",
                     "saida", "saída", "termination_date"],
    "email":        ["email", "e-mail", "mail", "email_corporativo"],
    "telefone":     ["telefone", "tel", "phone", "celular", "ramal"],
    "status":       ["status", "situação", "situacao", "ativo",
                     "active", "tipo_contrato"],
}


def _detectar_coluna(df: pd.DataFrame, grupo: str) -> str | None:
    candidatos = MAPA_COLUNAS.get(grupo, [])
    colunas_lower = {c.lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in colunas_lower:
            return colunas_lower[cand.lower()]
    return None


def _tempo_empresa(data_admissao: str) -> str:
    """Calcula tempo de empresa a partir da data de admissão."""
    if not data_admissao or str(data_admissao).strip() in ("nan", "None", ""):
        return ""
    try:
        dt = datetime.strptime(str(data_admissao), "%d/%m/%Y")
        hoje = datetime.today()
        delta = hoje - dt
        anos = delta.days // 365
        meses = (delta.days % 365) // 30
        if anos > 0:
            return f"{anos} ano{'s' if anos > 1 else ''} e {meses} mês{'es' if meses > 1 else ''}"
        return f"{meses} mês{'es' if meses > 1 else ''}"
    except Exception:
        return ""


def _faixa_salarial(salario) -> str:
    """Classifica faixa salarial baseada no salário mínimo BR (R$1.412)."""
    SM = 1412.0  # Salário mínimo 2024
    try:
        v = float(str(salario).replace("R$","").replace(".","").replace(",",".").strip())
    except Exception:
        return "Não informado"

    if v <= 0:
        return "Não informado"
    if v <= SM:
        return "Até 1 SM"
    if v <= SM * 2:
        return "1 a 2 SM"
    if v <= SM * 4:
        return "2 a 4 SM"
    if v <= SM * 8:
        return "4 a 8 SM"
    return "Acima de 8 SM"


def organizar_funcionarios(caminho_entrada: str, caminho_saida: str) -> dict:
    ext = caminho_entrada.lower().split(".")[-1]
    if ext == "csv":
        df = pd.read_csv(caminho_entrada, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(caminho_entrada)

    if df.empty:
        raise ValueError("A planilha está vazia ou não contém dados válidos.")

    df, stats = limpeza_basica(df)

    col_nome   = _detectar_coluna(df, "nome")
    col_cpf    = _detectar_coluna(df, "cpf")
    col_cargo  = _detectar_coluna(df, "cargo")
    col_dept   = _detectar_coluna(df, "departamento")
    col_sal    = _detectar_coluna(df, "salario")
    col_adm    = _detectar_coluna(df, "admissao")
    col_dem    = _detectar_coluna(df, "demissao")
    col_email  = _detectar_coluna(df, "email")
    col_tel    = _detectar_coluna(df, "telefone")
    col_status = _detectar_coluna(df, "status")

    # ── Padroniza NOME ────────────────────────────────────
    if col_nome:
        df[col_nome] = df[col_nome].apply(titulo_proprio)

    # ── Formata CPF ───────────────────────────────────────
    if col_cpf:
        df[col_cpf] = df[col_cpf].apply(formatar_cpf)

    # ── Padroniza CARGO e DEPARTAMENTO ───────────────────
    if col_cargo:
        df[col_cargo] = df[col_cargo].apply(titulo_proprio)
    if col_dept:
        df[col_dept] = df[col_dept].apply(titulo_proprio)

    # ── Padroniza STATUS ──────────────────────────────────
    if col_status:
        def normalizar_status_func(v):
            if not v or str(v).strip() in ("nan", "None", ""):
                return "Não informado"
            lower = str(v).strip().lower()
            mapa = {
                "ativo": "🟢 Ativo", "active": "🟢 Ativo",
                "inativo": "🔴 Inativo", "inactive": "🔴 Inativo",
                "afastado": "🟡 Afastado", "ferias": "🔵 Férias",
                "férias": "🔵 Férias", "demitido": "⚫ Demitido",
                "clt": "CLT", "pj": "PJ", "estágio": "Estágio",
                "estagio": "Estágio",
            }
            return mapa.get(lower, titulo_proprio(str(v)))
        df[col_status] = df[col_status].apply(normalizar_status_func)

    # ── Formata SALÁRIO ───────────────────────────────────
    if col_sal:
        sal_num = df[col_sal].apply(limpar_valor_monetario)

        # Cria FAIXA SALARIAL
        col_faixa = "Faixa Salarial"
        stats["colunas_criadas"].append(col_faixa)
        df[col_faixa] = sal_num.apply(_faixa_salarial)

        # Formata como moeda BR
        df[col_sal] = sal_num.apply(
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notna(v) else ""
        )

    # ── Padroniza DATAS ───────────────────────────────────
    if col_adm:
        df[col_adm] = df[col_adm].apply(padronizar_data)

        # Cria coluna TEMPO DE EMPRESA
        col_tempo = "Tempo de Empresa"
        stats["colunas_criadas"].append(col_tempo)
        df[col_tempo] = df[col_adm].apply(_tempo_empresa)

    if col_dem:
        df[col_dem] = df[col_dem].apply(padronizar_data)

    # ── Padroniza EMAIL e TELEFONE ────────────────────────
    if col_email:
        df[col_email] = df[col_email].apply(
            lambda x: str(x).strip().lower() if pd.notna(x) and str(x).strip() not in ("nan","None","") else ""
        )
    if col_tel:
        df[col_tel] = df[col_tel].apply(formatar_telefone)

    # ── Ordena por DEPARTAMENTO e NOME ───────────────────
    sort_cols = []
    if col_dept:
        sort_cols.append(col_dept)
    if col_nome:
        sort_cols.append(col_nome)
    if sort_cols:
        df = df.sort_values(sort_cols, na_position="last")

    df = df.reset_index(drop=True)
    stats["linhas_finais"] = len(df)

    exportar_xlsx_estilizado(df, caminho_saida, nome_aba="Funcionários")
    return stats
