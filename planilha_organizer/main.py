"""
╔══════════════════════════════════════════════════════════════╗
║          PLANILHA ORGANIZER PRO  —  v2.0                     ║
║          Software comercial de organização de planilhas      ║
╚══════════════════════════════════════════════════════════════╝
"""

import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path

from modules.financeira   import organizar_financeira
from modules.estoque      import organizar_estoque
from modules.vendas       import organizar_vendas
from modules.leads        import organizar_leads
from modules.funcionarios import organizar_funcionarios
from modules.generica     import organizar_generica
from utils.helpers        import validar_arquivo, gerar_nome_saida

# ── Tema global ──────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ════════════════════════════════════════════════════════════
#  PALETA  —  design system completo
# ════════════════════════════════════════════════════════════
C = {
    # Backgrounds
    "bg":           "#0B0D13",
    "sidebar":      "#0F1219",
    "panel":        "#13161F",
    "card":         "#181C28",
    "card_hover":   "#1E2335",
    "card_sel":     "#1A2540",
    "input_bg":     "#0E1018",
    "divider":      "#1F2437",
    # Accent — azul índigo
    "accent":       "#4B7BF5",
    "accent_h":     "#3A6AE0",
    "accent_soft":  "#1E2F5E",
    # Semânticas
    "success":      "#22C55E",
    "success_soft": "#14311F",
    "warning":      "#F59E0B",
    "warning_soft": "#2D230A",
    "error":        "#EF4444",
    "error_soft":   "#2D1010",
    # Textos
    "t1":           "#F1F3F9",
    "t2":           "#8892AA",
    "t3":           "#4D5670",
    # Bordas
    "border":       "#1E2437",
    "border_sel":   "#4B7BF5",
}

# ── Tipos de planilha ────────────────────────────────────────
TIPOS = [
    {
        "id":    "financeira",
        "nome":  "Financeira",
        "sub":   "Receitas · Despesas · Fluxo de Caixa",
        "icone": "💰",
        "cor":   "#22C55E",
        "func":  organizar_financeira,
        "dicas": ["Formata valores em R$", "Ordena por data", "Detecta receita/despesa"],
    },
    {
        "id":    "estoque",
        "nome":  "Estoque",
        "sub":   "Produtos · Quantidades · Preços",
        "icone": "📦",
        "cor":   "#F59E0B",
        "func":  organizar_estoque,
        "dicas": ["Cria status de estoque", "Calcula valor total", "Padroniza SKU/código"],
    },
    {
        "id":    "vendas",
        "nome":  "Vendas",
        "sub":   "Pedidos · Clientes · Faturamento",
        "icone": "📈",
        "cor":   "#4B7BF5",
        "func":  organizar_vendas,
        "dicas": ["Normaliza status dos pedidos", "Cria coluna Mês/Ano", "Ordena por data recente"],
    },
    {
        "id":    "leads",
        "nome":  "Leads & Clientes",
        "sub":   "Contatos · E-mails · Telefones",
        "icone": "👥",
        "cor":   "#A855F7",
        "func":  organizar_leads,
        "dicas": ["Remove e-mails duplicados", "Formata telefones BR", "Valida e-mails"],
    },
    {
        "id":    "funcionarios",
        "nome":  "Funcionários",
        "sub":   "RH · Salários · Departamentos",
        "icone": "🏢",
        "cor":   "#EC4899",
        "func":  organizar_funcionarios,
        "dicas": ["Calcula tempo de empresa", "Classifica faixa salarial", "Formata CPF"],
    },
    {
        "id":    "generica",
        "nome":  "Genérica",
        "sub":   "Qualquer planilha · Limpeza geral",
        "icone": "✨",
        "cor":   "#06B6D4",
        "func":  organizar_generica,
        "dicas": ["Detecta tipo de cada coluna", "Remove linhas vazias", "Padroniza textos"],
    },
]


# ════════════════════════════════════════════════════════════
#  WIDGET — Card de tipo de planilha
# ════════════════════════════════════════════════════════════
class CardTipo(ctk.CTkFrame):
    def __init__(self, master, tipo: dict, callback, **kwargs):
        super().__init__(
            master,
            fg_color=C["card"],
            corner_radius=14,
            border_width=1,
            border_color=C["border"],
            **kwargs,
        )
        self.tipo      = tipo
        self.callback  = callback
        self._selected = False
        self.configure(cursor="hand2")

        self._dot = ctk.CTkLabel(
            self, text="●",
            font=ctk.CTkFont(size=10),
            text_color=C["divider"],
            width=14,
        )
        self._dot.place(x=10, y=10)

        self._ico = ctk.CTkLabel(
            self, text=tipo["icone"],
            font=ctk.CTkFont(size=30),
        )
        self._ico.pack(pady=(22, 6))

        self._nome = ctk.CTkLabel(
            self, text=tipo["nome"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=C["t1"],
        )
        self._nome.pack()

        self._sub = ctk.CTkLabel(
            self, text=tipo["sub"],
            font=ctk.CTkFont(size=10),
            text_color=C["t3"],
            wraplength=148,
            justify="center",
        )
        self._sub.pack(pady=(3, 18))

        for w in [self, self._ico, self._nome, self._sub, self._dot]:
            w.bind("<Button-1>", self._click)
            w.bind("<Enter>",    self._enter)
            w.bind("<Leave>",    self._leave)

    def _click(self, _=None):  self.callback(self.tipo["id"])

    def _enter(self, _=None):
        if not self._selected:
            self.configure(fg_color=C["card_hover"])
            self._dot.configure(text_color=self.tipo["cor"])

    def _leave(self, _=None):
        if not self._selected:
            self.configure(fg_color=C["card"])
            self._dot.configure(text_color=C["divider"])

    def selecionar(self):
        self._selected = True
        self.configure(fg_color=C["card_sel"], border_color=self.tipo["cor"])
        self._nome.configure(text_color=self.tipo["cor"])
        self._dot.configure(text_color=self.tipo["cor"])

    def desselecionar(self):
        self._selected = False
        self.configure(fg_color=C["card"], border_color=C["border"])
        self._nome.configure(text_color=C["t1"])
        self._dot.configure(text_color=C["divider"])


# ════════════════════════════════════════════════════════════
#  WIDGET — Toast / notificação flutuante
# ════════════════════════════════════════════════════════════
class Toast(ctk.CTkToplevel):
    def __init__(self, master, titulo: str, msg: str, tipo: str = "sucesso"):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes("-topmost", True)

        cor_map = {
            "sucesso": C["success"],
            "erro":    C["error"],
            "aviso":   C["warning"],
            "info":    C["accent"],
        }
        ico_map = {"sucesso": "✓", "erro": "✕", "aviso": "!", "info": "i"}
        cor = cor_map.get(tipo, C["accent"])
        ico = ico_map.get(tipo, "i")

        self.configure(fg_color=C["panel"])

        outer = ctk.CTkFrame(self, fg_color=cor, corner_radius=12)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color=C["panel"], corner_radius=10)
        inner.pack(fill="both", expand=True, padx=(4, 0), pady=0)
        content = ctk.CTkFrame(inner, fg_color="transparent")
        content.pack(padx=18, pady=16, fill="both")

        ico_f = ctk.CTkFrame(content, width=36, height=36,
                             fg_color=cor + "22", corner_radius=18)
        ico_f.pack(side="left", padx=(0, 12))
        ico_f.pack_propagate(False)
        ctk.CTkLabel(ico_f, text=ico,
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=cor).pack(expand=True)

        tf = ctk.CTkFrame(content, fg_color="transparent")
        tf.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(tf, text=titulo,
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=C["t1"], anchor="w").pack(anchor="w")
        ctk.CTkLabel(tf, text=msg,
                     font=ctk.CTkFont(size=11),
                     text_color=C["t2"], anchor="w",
                     wraplength=260).pack(anchor="w", pady=(2, 0))

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 340, 88
        self.geometry(f"{w}x{h}+{sw-w-24}+{sh-h-60}")
        self.after(4000, self._fechar)

    def _fechar(self):
        try: self.destroy()
        except Exception: pass


# ════════════════════════════════════════════════════════════
#  WIDGET — Diálogo de resultado customizado
# ════════════════════════════════════════════════════════════
class ResultadoDialog(ctk.CTkToplevel):
    def __init__(self, master, resultado: dict, caminho_saida: str):
        super().__init__(master)
        self.title("Resultado")
        self.geometry("520x460")
        self.resizable(False, False)
        self.configure(fg_color=C["panel"])
        self.attributes("-topmost", True)
        self.grab_set()

        self.update_idletasks()
        px = master.winfo_x() + master.winfo_width()  // 2 - 260
        py = master.winfo_y() + master.winfo_height() // 2 - 230
        self.geometry(f"520x460+{px}+{py}")
        self._build(resultado, caminho_saida)

    def _build(self, r: dict, caminho: str):
        nome = Path(caminho).name
        pasta = str(Path(caminho).parent)

        # ── Banner de sucesso ────────────────────────────
        banner = ctk.CTkFrame(self, fg_color=C["success_soft"], corner_radius=0)
        banner.pack(fill="x")
        ctk.CTkLabel(banner, text="✓",
                     font=ctk.CTkFont(size=40, weight="bold"),
                     text_color=C["success"]).pack(pady=(26, 4))
        ctk.CTkLabel(banner, text="Planilha organizada com sucesso!",
                     font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                     text_color=C["t1"]).pack()
        ctk.CTkLabel(banner, text="Seu arquivo foi processado e está pronto para uso.",
                     font=ctk.CTkFont(size=11),
                     text_color=C["t2"]).pack(pady=(4, 22))

        # ── Métricas ─────────────────────────────────────
        met = ctk.CTkFrame(self, fg_color="transparent")
        met.pack(fill="x", padx=24, pady=(20, 0))
        dados = [
            ("Linhas\nprocessadas",  str(r.get("linhas_originais", "—")), C["accent"]),
            ("Duplicados\nremovidos", str(r.get("duplicados", 0)),         C["warning"]),
            ("Resultado\nfinal",      str(r.get("linhas_finais", "—")),    C["success"]),
        ]
        for i, (lbl, val, cor) in enumerate(dados):
            met.columnconfigure(i, weight=1)
            b = ctk.CTkFrame(met, fg_color=C["card"], corner_radius=10)
            b.grid(row=0, column=i, padx=4, sticky="ew")
            ctk.CTkLabel(b, text=val,
                         font=ctk.CTkFont(size=24, weight="bold"),
                         text_color=cor).pack(pady=(14, 2))
            ctk.CTkLabel(b, text=lbl,
                         font=ctk.CTkFont(size=10),
                         text_color=C["t3"],
                         justify="center").pack(pady=(0, 14))

        # ── Arquivo gerado ────────────────────────────────
        af = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=10)
        af.pack(fill="x", padx=24, pady=(14, 0))
        ai = ctk.CTkFrame(af, fg_color="transparent")
        ai.pack(padx=16, pady=14, fill="x")
        ctk.CTkLabel(ai, text="📄",
                     font=ctk.CTkFont(size=24)).pack(side="left", padx=(0, 12))
        tf = ctk.CTkFrame(ai, fg_color="transparent")
        tf.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(tf, text=nome,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=C["t1"], anchor="w").pack(anchor="w")
        ctk.CTkLabel(tf, text=pasta,
                     font=ctk.CTkFont(size=10),
                     text_color=C["t3"], anchor="w").pack(anchor="w")

        # ── Colunas criadas ───────────────────────────────
        if r.get("colunas_criadas"):
            cols = ", ".join(r["colunas_criadas"])
            ex = ctk.CTkFrame(self, fg_color=C["accent_soft"], corner_radius=10)
            ex.pack(fill="x", padx=24, pady=(10, 0))
            ctk.CTkLabel(ex,
                         text=f"✦  Novas colunas criadas automaticamente:  {cols}",
                         font=ctk.CTkFont(size=11),
                         text_color=C["accent"]).pack(padx=14, pady=10)

        # ── Botões ────────────────────────────────────────
        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=24, pady=(16, 22))
        ctk.CTkButton(
            btns, text="📁  Abrir Pasta",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_h"],
            corner_radius=10, height=42,
            command=lambda: self._abrir(Path(caminho).parent),
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            btns, text="Fechar",
            font=ctk.CTkFont(size=13),
            fg_color=C["card"], hover_color=C["card_hover"],
            border_width=1, border_color=C["border"],
            text_color=C["t2"], corner_radius=10, height=42,
            command=self.destroy,
        ).pack(side="left", fill="x", expand=True, padx=(8, 0))

    def _abrir(self, pasta: Path):
        try:
            if sys.platform == "win32": os.startfile(str(pasta))
            elif sys.platform == "darwin": os.system(f'open "{pasta}"')
            else: os.system(f'xdg-open "{pasta}"')
        except Exception: pass


# ════════════════════════════════════════════════════════════
#  APLICAÇÃO PRINCIPAL
# ════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Planilha Organizer Pro")
        self.geometry("1040x700")
        self.minsize(920, 620)
        self.configure(fg_color=C["bg"])
        self._centralizar()

        self._arquivo: str | None = None
        self._tipo:    str | None = None
        self._cards:   dict[str, CardTipo] = {}

        self._build_layout()

    def _centralizar(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 1040, 700
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ════════════════════════════════════════════════════
    #  LAYOUT:  sidebar  |  área principal
    # ════════════════════════════════════════════════════
    def _build_layout(self):
        self._sidebar = ctk.CTkFrame(self, fg_color=C["sidebar"],
                                     width=264, corner_radius=0)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)

        self._main = ctk.CTkScrollableFrame(
            self, fg_color=C["bg"],
            scrollbar_fg_color=C["sidebar"],
            scrollbar_button_color=C["divider"],
            scrollbar_button_hover_color=C["accent"],
            corner_radius=0,
        )
        self._main.pack(side="left", fill="both", expand=True)

        self._build_sidebar()
        self._build_main()

    # ════════════════════════════════════════════════════
    #  SIDEBAR
    # ════════════════════════════════════════════════════
    def _build_sidebar(self):
        sb = self._sidebar

        # ── Logo ─────────────────────────────────────────
        logo = ctk.CTkFrame(sb, fg_color="transparent")
        logo.pack(padx=22, pady=(30, 0), anchor="w")

        badge = ctk.CTkFrame(logo, fg_color=C["accent_soft"],
                             corner_radius=10, width=40, height=40)
        badge.pack(side="left", padx=(0, 12))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text="◈",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=C["accent"]).pack(expand=True)

        lt = ctk.CTkFrame(logo, fg_color="transparent")
        lt.pack(side="left")
        ctk.CTkLabel(lt, text="Planilha Organizer",
                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                     text_color=C["t1"]).pack(anchor="w")
        ctk.CTkLabel(lt, text="Pro  v2.0",
                     font=ctk.CTkFont(size=11),
                     text_color=C["accent"]).pack(anchor="w")

        # ── Divisor ───────────────────────────────────────
        ctk.CTkFrame(sb, fg_color=C["divider"], height=1).pack(
            fill="x", padx=20, pady=(24, 18))

        # ── Nav ──────────────────────────────────────────
        ctk.CTkLabel(sb, text="MENU",
                     font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=C["t3"]).pack(anchor="w", padx=22, pady=(0, 6))

        nav_items = [
            ("Organizar Planilha", "⚡", True),
            ("Histórico",          "📋", False),
            ("Configurações",      "⚙",  False),
            ("Ajuda",              "?",  False),
        ]
        for lbl, ico, ativo in nav_items:
            bg   = C["accent_soft"] if ativo else "transparent"
            tcor = C["accent"] if ativo else C["t2"]
            f = ctk.CTkFrame(sb, fg_color=bg, corner_radius=8)
            f.pack(fill="x", padx=14, pady=1)
            row = ctk.CTkFrame(f, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=9)
            ctk.CTkLabel(row, text=ico,
                         font=ctk.CTkFont(size=13),
                         text_color=tcor, width=20).pack(side="left")
            ctk.CTkLabel(row, text=lbl,
                         font=ctk.CTkFont(size=12,
                             weight="bold" if ativo else "normal"),
                         text_color=tcor, anchor="w").pack(side="left", padx=8)

        # ── Divisor ───────────────────────────────────────
        ctk.CTkFrame(sb, fg_color=C["divider"], height=1).pack(
            fill="x", padx=20, pady=(18, 14))

        # ── Painel de dicas ───────────────────────────────
        ctk.CTkLabel(sb, text="AÇÕES DO TIPO SELECIONADO",
                     font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=C["t3"]).pack(anchor="w", padx=22, pady=(0, 8))

        self._dicas_frame = ctk.CTkFrame(sb, fg_color="transparent")
        self._dicas_frame.pack(fill="x", padx=14)
        self._render_dicas(None)

        # ── Rodapé ────────────────────────────────────────
        ctk.CTkFrame(sb, fg_color=C["divider"], height=1).pack(
            side="bottom", fill="x", padx=20, pady=(0, 12))
        ctk.CTkLabel(sb, text="© 2026  Planilha Organizer Pro",
                     font=ctk.CTkFont(size=9),
                     text_color=C["t3"]).pack(side="bottom", pady=(0, 4))

    def _render_dicas(self, tipo_id):
        for w in self._dicas_frame.winfo_children():
            w.destroy()
        if tipo_id is None:
            ctk.CTkLabel(self._dicas_frame,
                         text="Selecione um tipo de\nplanilha →",
                         font=ctk.CTkFont(size=11),
                         text_color=C["t3"],
                         justify="left").pack(anchor="w", padx=8, pady=4)
            return
        tipo = next((t for t in TIPOS if t["id"] == tipo_id), None)
        if not tipo: return
        for dica in tipo["dicas"]:
            row = ctk.CTkFrame(self._dicas_frame, fg_color="transparent")
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text="→",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=tipo["cor"], width=16).pack(side="left")
            ctk.CTkLabel(row, text=dica,
                         font=ctk.CTkFont(size=11),
                         text_color=C["t2"], anchor="w").pack(side="left", padx=6)

    # ════════════════════════════════════════════════════
    #  ÁREA PRINCIPAL
    # ════════════════════════════════════════════════════
    def _build_main(self):
        P = {"padx": 32}

        # ── Cabeçalho de página ──────────────────────────
        hdr = ctk.CTkFrame(self._main, fg_color="transparent")
        hdr.pack(fill="x", **P, pady=(32, 0))
        ctk.CTkLabel(hdr, text="Organizar Planilha",
                     font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"),
                     text_color=C["t1"], anchor="w").pack(anchor="w")
        ctk.CTkLabel(hdr,
                     text="Selecione o arquivo, escolha o tipo e clique em organizar.",
                     font=ctk.CTkFont(size=13),
                     text_color=C["t2"], anchor="w").pack(anchor="w", pady=(4, 0))

        ctk.CTkFrame(self._main, fg_color=C["divider"], height=1).pack(
            fill="x", padx=32, pady=(18, 0))

        # ── Passo 1 ───────────────────────────────────────
        self._step_header("01", "Selecione o arquivo",
                          "Formatos aceitos: .xlsx  ·  .csv  (até 50 MB)")
        self._build_secao_arquivo()

        # ── Passo 2 ───────────────────────────────────────
        self._step_header("02", "Tipo de planilha",
                          "Não tem certeza? Use Genérica — funciona para qualquer planilha.")
        self._build_secao_tipo()

        # ── Botão ─────────────────────────────────────────
        self._build_botao()

        # ── Status + progresso ────────────────────────────
        self._build_status()

    def _step_header(self, num: str, titulo: str, sub: str):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.pack(fill="x", padx=32, pady=(26, 10))

        badge = ctk.CTkFrame(f, fg_color=C["accent_soft"],
                             corner_radius=8, width=38, height=38)
        badge.pack(side="left", padx=(0, 14))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text=num,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=C["accent"]).pack(expand=True)

        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.pack(side="left")
        ctk.CTkLabel(tf, text=titulo,
                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                     text_color=C["t1"]).pack(anchor="w")
        ctk.CTkLabel(tf, text=sub,
                     font=ctk.CTkFont(size=11),
                     text_color=C["t3"]).pack(anchor="w")

    # ── Seção de arquivo ─────────────────────────────────
    def _build_secao_arquivo(self):
        outer = ctk.CTkFrame(self._main, fg_color=C["card"],
                             corner_radius=14, border_width=1,
                             border_color=C["border"])
        outer.pack(fill="x", padx=32)

        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)

        # Upload icon
        up = ctk.CTkFrame(inner, fg_color=C["accent_soft"],
                          corner_radius=12, width=52, height=52)
        up.pack(side="left", padx=(0, 16))
        up.pack_propagate(False)
        ctk.CTkLabel(up, text="↑",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=C["accent"]).pack(expand=True)

        txt = ctk.CTkFrame(inner, fg_color="transparent")
        txt.pack(side="left", fill="x", expand=True)

        self._lbl_filename = ctk.CTkLabel(
            txt, text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=C["t3"], anchor="w")
        self._lbl_filename.pack(anchor="w")

        self._lbl_filemeta = ctk.CTkLabel(
            txt, text="Clique em 'Escolher Arquivo' para começar",
            font=ctk.CTkFont(size=11),
            text_color=C["t3"], anchor="w")
        self._lbl_filemeta.pack(anchor="w", pady=(2, 0))

        ctk.CTkButton(
            inner, text="Escolher Arquivo",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_h"],
            corner_radius=10, height=42, width=160,
            command=self._selecionar_arquivo,
        ).pack(side="right")

    # ── Seção de tipos ────────────────────────────────────
    def _build_secao_tipo(self):
        grid = ctk.CTkFrame(self._main, fg_color="transparent")
        grid.pack(fill="x", padx=32)
        for col in range(3):
            grid.columnconfigure(col, weight=1)
        for idx, tipo in enumerate(TIPOS):
            r, c = divmod(idx, 3)
            card = CardTipo(grid, tipo=tipo,
                            callback=self._selecionar_tipo, width=180)
            card.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
            self._cards[tipo["id"]] = card

    # ── Botão principal ───────────────────────────────────
    def _build_botao(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.pack(fill="x", padx=32, pady=(22, 0))
        self._btn = ctk.CTkButton(
            f, text="⚡   Organizar Planilha",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_h"],
            corner_radius=12, height=56,
            command=self._iniciar,
        )
        self._btn.pack(fill="x")

    # ── Status + progresso ────────────────────────────────
    def _build_status(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.pack(fill="x", padx=32, pady=(10, 0))

        sf = ctk.CTkFrame(f, fg_color=C["card"],
                          corner_radius=8, border_width=1,
                          border_color=C["border"])
        sf.pack(fill="x")
        si = ctk.CTkFrame(sf, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self._sdot = ctk.CTkLabel(si, text="●",
                                  font=ctk.CTkFont(size=10),
                                  text_color=C["t3"], width=14)
        self._sdot.pack(side="left", padx=(0, 8))

        self._slbl = ctk.CTkLabel(
            si,
            text="Aguardando — selecione o arquivo e o tipo da planilha.",
            font=ctk.CTkFont(size=12),
            text_color=C["t2"], anchor="w")
        self._slbl.pack(side="left", fill="x", expand=True)

        self._prog = ctk.CTkProgressBar(
            f, fg_color=C["divider"],
            progress_color=C["accent"],
            corner_radius=4, height=4)
        self._prog.pack(fill="x", pady=(8, 36))
        self._prog.set(0)

    # ════════════════════════════════════════════════════
    #  HELPERS UI
    # ════════════════════════════════════════════════════
    def _set_status(self, msg: str, tipo: str = "info"):
        cor = {"info":C["t3"],"sucesso":C["success"],
               "erro":C["error"],"aviso":C["warning"],
               "loading":C["accent"]}.get(tipo, C["t3"])
        self._sdot.configure(text_color=cor)
        self._slbl.configure(text=msg)

    # ════════════════════════════════════════════════════
    #  AÇÕES
    # ════════════════════════════════════════════════════
    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione sua planilha",
            filetypes=[("Planilhas Excel","*.xlsx"),
                       ("Arquivos CSV","*.csv"),
                       ("Todos","*.*")])
        if not caminho: return

        ok, msg = validar_arquivo(caminho)
        if not ok:
            self._set_status(msg, "erro")
            Toast(self, "Arquivo inválido", msg, "erro")
            return

        self._arquivo = caminho
        path = Path(caminho)
        size = path.stat().st_size
        size_str = (f"{size} B" if size < 1024 else
                    f"{size/1024:.1f} KB" if size < 1048576 else
                    f"{size/1048576:.1f} MB")
        nome_exib = path.name if len(path.name) <= 55 else f"…{path.name[-52:]}"

        self._lbl_filename.configure(text=nome_exib, text_color=C["t1"])
        self._lbl_filemeta.configure(
            text=f"{path.suffix.upper()}  ·  {size_str}  ·  {path.parent}",
            text_color=C["t3"])
        self._set_status(f"Arquivo carregado: {path.name}", "sucesso")
        self._prog.set(0.12)

    def _selecionar_tipo(self, tipo_id: str):
        for card in self._cards.values():
            card.desselecionar()
        self._tipo = tipo_id
        self._cards[tipo_id].selecionar()
        self._render_dicas(tipo_id)
        nome = next(t["nome"] for t in TIPOS if t["id"] == tipo_id)
        self._set_status(f"Tipo selecionado: {nome}", "info")
        if self._arquivo:
            self._prog.set(0.30)

    def _iniciar(self):
        if not self._arquivo:
            Toast(self, "Arquivo não selecionado",
                  "Escolha um arquivo .xlsx ou .csv primeiro.", "aviso")
            return
        if not self._tipo:
            Toast(self, "Tipo não selecionado",
                  "Clique em um dos cards para escolher o tipo.", "aviso")
            return

        self._btn.configure(state="disabled", text="  Processando...")
        self._prog.set(0)
        self._prog.configure(mode="indeterminate")
        self._prog.start(10)
        self._set_status("Organizando sua planilha, por favor aguarde...", "loading")
        threading.Thread(target=self._executar, daemon=True).start()

    def _executar(self):
        try:
            func   = next(t["func"] for t in TIPOS if t["id"] == self._tipo)
            saida  = gerar_nome_saida(self._arquivo, self._tipo)
            result = func(self._arquivo, saida)
            self.after(0, self._on_sucesso, saida, result)
        except Exception as e:
            self.after(0, self._on_erro, str(e))

    def _on_sucesso(self, saida: str, r: dict):
        self._prog.stop()
        self._prog.configure(mode="determinate", progress_color=C["success"])
        self._prog.set(1.0)
        self._btn.configure(state="normal", text="⚡   Organizar Planilha")
        self._set_status(
            f"Concluído!  {r.get('linhas_finais','?')} linhas salvas em {Path(saida).name}",
            "sucesso")
        Toast(self, "Pronto!",
              f"{r.get('linhas_finais','?')} linhas organizadas.", "sucesso")
        self.after(300, lambda: ResultadoDialog(self, r, saida))
        self.after(4000, lambda: self._prog.configure(progress_color=C["accent"]))

    def _on_erro(self, msg: str):
        self._prog.stop()
        self._prog.configure(mode="determinate", progress_color=C["error"])
        self._prog.set(0)
        self._btn.configure(state="normal", text="⚡   Organizar Planilha")
        self._set_status(f"Erro: {msg[:90]}", "erro")
        Toast(self, "Erro ao processar", msg[:120], "erro")
        self.after(3000, lambda: self._prog.configure(progress_color=C["accent"]))


# ════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
