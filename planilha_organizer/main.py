"""
╔══════════════════════════════════════════════════════════════╗
║          PLANILHA ORGANIZER PRO  —  v2.0                     ║
║          Software comercial de organização de planilhas      ║
╚══════════════════════════════════════════════════════════════╝
"""

import os
import sys
import threading
import multiprocessing
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk

# ─────────────────────────────────────────────────────────────
#  BASE DIR — funciona em modo normal e em .EXE (PyInstaller)
# ─────────────────────────────────────────────────────────────
BASE_DIR = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
sys.path.append(str(BASE_DIR))

# ─────────────────────────────────────────────────────────────
#  IMPORTS DO PROJETO
# ─────────────────────────────────────────────────────────────
from modules.financeira   import organizar_financeira
from modules.estoque      import organizar_estoque
from modules.vendas       import organizar_vendas
from modules.leads        import organizar_leads
from modules.funcionarios import organizar_funcionarios
from modules.generica     import organizar_generica
from utils.helpers        import validar_arquivo, gerar_nome_saida

# ─────────────────────────────────────────────────────────────
#  TEMA GLOBAL
# ─────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ════════════════════════════════════════════════════════════
#  PALETA DE CORES — Luxury Dark
# ════════════════════════════════════════════════════════════
C = {
    "bg":           "#080A10",
    "bg2":          "#0C0E18",
    "sidebar":      "#0A0C14",
    "panel":        "#0F1220",
    "card":         "#111526",
    "card_hover":   "#161A2E",
    "card_sel":     "#1A2040",
    "card_border":  "#1E2438",
    "card_sel_brd": "#3D6FE8",
    "input_bg":     "#0A0C18",
    "divider":      "#181C30",
    "accent":       "#3D6FE8",
    "accent_h":     "#2F5ED0",
    "accent_glow":  "#1E3580",
    "accent_soft":  "#1A2550",
    "gold":         "#E8B83D",
    "gold_soft":    "#2A1F05",
    "success":      "#22C55E",
    "success_soft": "#0C2018",
    "warning":      "#F59E0B",
    "warning_soft": "#221600",
    "error":        "#EF4444",
    "error_soft":   "#1E0808",
    "t1":           "#EDF0FA",
    "t2":           "#7A85A0",
    "t3":           "#3D4560",
    "t4":           "#252A40",
    "border":       "#1A1F35",
}

# ─────────────────────────────────────────────────────────────
#  TIPOS DE PLANILHA
# ─────────────────────────────────────────────────────────────
TIPOS = [
    {"id":"financeira", "nome":"Financeira",   "sub":"Receitas · Despesas · Fluxo",
     "icone":"💰", "cor":"#22C55E", "cor_soft":"#0A2010", "func":organizar_financeira,
     "dicas":["Formata valores", "Ordena datas", "Detecta receita/despesa"]},

    {"id":"estoque",    "nome":"Estoque",      "sub":"Produtos · Quantidades",
     "icone":"📦", "cor":"#F59E0B", "cor_soft":"#1C1200", "func":organizar_estoque,
     "dicas":["Calcula estoque", "Cria alertas", "Padroniza códigos"]},

    {"id":"vendas",     "nome":"Vendas",       "sub":"Pedidos · Clientes",
     "icone":"📈", "cor":"#3D6FE8", "cor_soft":"#101830", "func":organizar_vendas,
     "dicas":["Normaliza status", "Cria mês/ano", "Ordena vendas"]},

    {"id":"leads",      "nome":"Leads",        "sub":"E-mails · Telefones",
     "icone":"👥", "cor":"#A855F7", "cor_soft":"#180E28", "func":organizar_leads,
     "dicas":["Remove duplicados", "Valida e-mails", "Formata telefones"]},

    {"id":"funcionarios","nome":"Funcionários","sub":"RH · Salários",
     "icone":"🏢", "cor":"#EC4899", "cor_soft":"#1E0818", "func":organizar_funcionarios,
     "dicas":["Faixa salarial", "Tempo empresa", "Formata CPF"]},

    {"id":"generica",   "nome":"Genérica",     "sub":"Qualquer planilha",
     "icone":"✨", "cor":"#06B6D4", "cor_soft":"#021820", "func":organizar_generica,
     "dicas":["Remove vazios", "Detecta colunas", "Padroniza texto"]},
]


# ════════════════════════════════════════════════════════════
#  WIDGET: TipoCard — card selecionável premium
# ════════════════════════════════════════════════════════════
class TipoCard(ctk.CTkFrame):
    def __init__(self, master, tipo_data, on_select, **kwargs):
        super().__init__(
            master,
            fg_color=C["card"],
            border_color=C["card_border"],
            border_width=1,
            corner_radius=14,
            **kwargs
        )
        self.tipo_data = tipo_data
        self.on_select = on_select
        self.selected = False
        self._build()
        self._bind_hover()

    def _build(self):
        t = self.tipo_data

        # Ícone + badge colorido
        icon_frame = ctk.CTkFrame(
            self, fg_color=t["cor_soft"],
            corner_radius=10,
            width=48, height=48
        )
        icon_frame.pack(side="left", padx=(16, 12), pady=14)
        icon_frame.pack_propagate(False)

        icon_lbl = ctk.CTkLabel(
            icon_frame, text=t["icone"],
            font=ctk.CTkFont(size=22),
            text_color=t["cor"]
        )
        icon_lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Texto
        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.pack(side="left", fill="both", expand=True, pady=10)

        ctk.CTkLabel(
            text_frame, text=t["nome"],
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=C["t1"], anchor="w"
        ).pack(anchor="w")

        ctk.CTkLabel(
            text_frame, text=t["sub"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=C["t2"], anchor="w"
        ).pack(anchor="w")

        # Dicas — pequenas tags
        tags_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        tags_frame.pack(anchor="w", pady=(4, 0))
        for dica in t["dicas"]:
            tag = ctk.CTkFrame(
                tags_frame, fg_color=C["card_border"],
                corner_radius=6, height=18
            )
            tag.pack(side="left", padx=(0, 4))
            ctk.CTkLabel(
                tag, text=dica,
                font=ctk.CTkFont(size=9),
                text_color=C["t3"],
                padx=6, pady=1
            ).pack()

        # Checkmark à direita
        self.check_lbl = ctk.CTkLabel(
            self, text="○",
            font=ctk.CTkFont(size=18),
            text_color=C["t4"],
            width=36
        )
        self.check_lbl.pack(side="right", padx=14)

        # Bind em todos os widgets filhos
        for w in [self, icon_frame, icon_lbl, text_frame, tags_frame]:
            w.bind("<Button-1>", self._clicked)

    def _bind_hover(self):
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        for child in self.winfo_children():
            child.bind("<Enter>", self._on_enter)
            child.bind("<Leave>", self._on_leave)

    def _on_enter(self, _=None):
        if not self.selected:
            self.configure(fg_color=C["card_hover"])

    def _on_leave(self, _=None):
        if not self.selected:
            self.configure(fg_color=C["card"])

    def _clicked(self, _=None):
        self.on_select(self.tipo_data["id"])

    def set_selected(self, val: bool):
        self.selected = val
        t = self.tipo_data
        if val:
            self.configure(
                fg_color=C["card_sel"],
                border_color=t["cor"],
                border_width=2
            )
            self.check_lbl.configure(text="✓", text_color=t["cor"])
        else:
            self.configure(
                fg_color=C["card"],
                border_color=C["card_border"],
                border_width=1
            )
            self.check_lbl.configure(text="○", text_color=C["t4"])


# ════════════════════════════════════════════════════════════
#  WIDGET: FileDropZone — área de seleção de arquivo premium
# ════════════════════════════════════════════════════════════
class FileDropZone(ctk.CTkFrame):
    def __init__(self, master, on_file_selected, **kwargs):
        super().__init__(
            master,
            fg_color=C["input_bg"],
            border_color=C["border"],
            border_width=1,
            corner_radius=14,
            height=90,
            **kwargs
        )
        self.on_file_selected = on_file_selected
        self._file_path = None
        self._build()
        self.bind("<Button-1>", self._open_dialog)
        self.bind("<Enter>", self._hover_on)
        self.bind("<Leave>", self._hover_off)

    def _build(self):
        self.pack_propagate(False)

        self.inner = ctk.CTkFrame(self, fg_color="transparent")
        self.inner.place(relx=0.5, rely=0.5, anchor="center")

        self.icon_lbl = ctk.CTkLabel(
            self.inner, text="📂",
            font=ctk.CTkFont(size=26),
        )
        self.icon_lbl.pack()

        self.main_lbl = ctk.CTkLabel(
            self.inner,
            text="Clique para escolher o arquivo",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=C["t2"]
        )
        self.main_lbl.pack()

        self.sub_lbl = ctk.CTkLabel(
            self.inner,
            text=".xlsx  ·  .csv",
            font=ctk.CTkFont(size=10),
            text_color=C["t3"]
        )
        self.sub_lbl.pack()

        for w in [self.inner, self.icon_lbl, self.main_lbl, self.sub_lbl]:
            w.bind("<Button-1>", self._open_dialog)
            w.bind("<Enter>", self._hover_on)
            w.bind("<Leave>", self._hover_off)

    def _hover_on(self, _=None):
        self.configure(border_color=C["accent"], fg_color=C["accent_soft"])

    def _hover_off(self, _=None):
        if not self._file_path:
            self.configure(border_color=C["border"], fg_color=C["input_bg"])
        else:
            self.configure(border_color=C["success"], fg_color=C["success_soft"])

    def _open_dialog(self, _=None):
        path = filedialog.askopenfilename(
            filetypes=[("Planilhas", "*.xlsx *.csv")])
        if path:
            self.on_file_selected(path)

    def set_file(self, path: str, ok: bool):
        self._file_path = path if ok else None
        name = Path(path).name
        if ok:
            self.configure(border_color=C["success"], fg_color=C["success_soft"])
            self.icon_lbl.configure(text="✅")
            self.main_lbl.configure(text=name, text_color=C["success"])
            self.sub_lbl.configure(text="Arquivo carregado · clique para trocar")
        else:
            self.configure(border_color=C["error"], fg_color=C["error_soft"])
            self.icon_lbl.configure(text="❌")
            self.main_lbl.configure(text="Arquivo inválido", text_color=C["error"])
            self.sub_lbl.configure(text=name)

    def reset(self):
        self._file_path = None
        self.configure(border_color=C["border"], fg_color=C["input_bg"])
        self.icon_lbl.configure(text="📂")
        self.main_lbl.configure(
            text="Clique para escolher o arquivo", text_color=C["t2"])
        self.sub_lbl.configure(text=".xlsx  ·  .csv", text_color=C["t3"])


# ════════════════════════════════════════════════════════════
#  WIDGET: StatusBar — barra de status animada
# ════════════════════════════════════════════════════════════
class StatusBar(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(
            master,
            fg_color=C["panel"],
            corner_radius=12,
            border_color=C["border"],
            border_width=1,
            height=48,
            **kwargs
        )
        self.pack_propagate(False)
        self._dot = ctk.CTkLabel(
            self, text="●",
            font=ctk.CTkFont(size=10),
            text_color=C["t3"], width=22
        )
        self._dot.pack(side="left", padx=(14, 4))

        self._lbl = ctk.CTkLabel(
            self,
            text="Pronto — selecione um arquivo e tipo para começar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=C["t2"],
            anchor="w"
        )
        self._lbl.pack(side="left", fill="x", expand=True)

    def update(self, text, color_key="t2", dot_color=None):
        self._lbl.configure(text=text, text_color=C[color_key])
        self._dot.configure(text_color=dot_color or C[color_key])

    def idle(self):
        self.update("Pronto — selecione um arquivo e tipo para começar", "t3", C["t4"])

    def processing(self, text="Processando..."):
        self.update(text, "accent", C["accent"])

    def success(self, text):
        self.update(text, "success", C["success"])

    def error(self, text):
        self.update(text, "error", C["error"])

    def warning(self, text):
        self.update(text, "warning", C["warning"])


# ════════════════════════════════════════════════════════════
#  WIDGET: ProgressBar animada
# ════════════════════════════════════════════════════════════
class AnimatedProgress(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", height=3, **kwargs)
        self.pack_propagate(False)
        self._bar = ctk.CTkProgressBar(
            self,
            height=3,
            corner_radius=2,
            fg_color=C["panel"],
            progress_color=C["accent"]
        )
        self._bar.pack(fill="x", expand=True)
        self._bar.set(0)
        self._running = False
        self._val = 0

    def start(self):
        self._running = True
        self._val = 0
        self._bar.configure(mode="indeterminate")
        self._bar.start()

    def stop(self, success=True):
        self._running = False
        self._bar.stop()
        self._bar.configure(
            mode="determinate",
            progress_color=C["success"] if success else C["error"]
        )
        self._bar.set(1 if success else 0)

    def reset(self):
        self._bar.stop()
        self._bar.configure(mode="determinate", progress_color=C["accent"])
        self._bar.set(0)


# ════════════════════════════════════════════════════════════
#  APP PRINCIPAL — redesign premium
# ════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Planilha Organizer Pro")
        self.geometry("900x780")
        self.minsize(820, 680)
        self.configure(fg_color=C["bg"])
        self._arquivo  = None
        self._tipo     = None
        self._cards    = {}
        self._build_ui()

    # ─────────────────────────────────────────────────────────
    #  BUILD UI
    # ─────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Scroll container raiz
        self._root_scroll = ctk.CTkScrollableFrame(
            self,
            fg_color=C["bg"],
            scrollbar_button_color=C["t4"],
            scrollbar_button_hover_color=C["t3"],
        )
        self._root_scroll.pack(fill="both", expand=True)
        wrap = self._root_scroll

        # ─── HEADER ───────────────────────────────────────────
        header = ctk.CTkFrame(wrap, fg_color="transparent")
        header.pack(fill="x", padx=32, pady=(32, 0))

        # Badge "PRO"
        badge = ctk.CTkFrame(
            header, fg_color=C["gold_soft"],
            corner_radius=6, height=22, width=46
        )
        badge.pack(side="right", padx=(0, 0), pady=(6, 0))
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge, text="PRO",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=C["gold"]
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            header,
            text="Planilha Organizer",
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=C["t1"]
        ).pack(side="left", anchor="sw")

        # Linha decorativa sob o título
        divider = ctk.CTkFrame(wrap, fg_color=C["divider"], height=1)
        divider.pack(fill="x", padx=32, pady=(10, 0))

        # Subtítulo
        ctk.CTkLabel(
            wrap,
            text="Organize, limpe e padronize suas planilhas com precisão cirúrgica.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=C["t2"]
        ).pack(anchor="w", padx=32, pady=(8, 0))

        # ─── SEÇÃO: ARQUIVO ───────────────────────────────────
        self._section_label(wrap, "01", "Selecionar Arquivo")

        self._file_zone = FileDropZone(
            wrap, on_file_selected=self._selecionar_arquivo
        )
        self._file_zone.pack(fill="x", padx=32, pady=(0, 4))

        # ─── SEÇÃO: TIPO ──────────────────────────────────────
        self._section_label(wrap, "02", "Tipo de Planilha")

        cards_frame = ctk.CTkFrame(wrap, fg_color="transparent")
        cards_frame.pack(fill="x", padx=32)
        cards_frame.columnconfigure(0, weight=1)
        cards_frame.columnconfigure(1, weight=1)

        for i, t in enumerate(TIPOS):
            row, col = divmod(i, 2)
            card = TipoCard(
                cards_frame,
                tipo_data=t,
                on_select=self._selecionar_tipo
            )
            card.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            self._cards[t["id"]] = card

        # ─── PROGRESS BAR ─────────────────────────────────────
        self._progress = AnimatedProgress(wrap)
        self._progress.pack(fill="x", padx=32, pady=(20, 0))

        # ─── STATUS BAR ───────────────────────────────────────
        self._status = StatusBar(wrap)
        self._status.pack(fill="x", padx=32, pady=(8, 0))

        # ─── BOTÃO PRINCIPAL ──────────────────────────────────
        self._btn_run = ctk.CTkButton(
            wrap,
            text="  ⚡  Organizar Planilha",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=52,
            corner_radius=14,
            fg_color=C["accent"],
            hover_color=C["accent_h"],
            text_color=C["t1"],
            command=self._iniciar
        )
        self._btn_run.pack(fill="x", padx=32, pady=(12, 0))

        # ─── RODAPÉ ───────────────────────────────────────────
        footer = ctk.CTkFrame(wrap, fg_color="transparent")
        footer.pack(fill="x", padx=32, pady=(16, 32))

        ctk.CTkLabel(
            footer,
            text="Planilha Organizer Pro  ·  v2.0  ·  Todos os direitos reservados",
            font=ctk.CTkFont(size=10),
            text_color=C["t4"]
        ).pack()

    # ─────────────────────────────────────────────────────────
    #  HELPER: rótulo de seção numerado
    # ─────────────────────────────────────────────────────────
    def _section_label(self, parent, num: str, label: str):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=32, pady=(22, 8))

        badge = ctk.CTkFrame(
            row, fg_color=C["accent_soft"],
            corner_radius=6, width=26, height=20
        )
        badge.pack(side="left", padx=(0, 8))
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge, text=num,
            font=ctk.CTkFont(size=9, weight="bold"),
            text_color=C["accent"]
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            row, text=label,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=C["t2"]
        ).pack(side="left")

    # ─────────────────────────────────────────────────────────
    #  AÇÕES
    # ─────────────────────────────────────────────────────────
    def _selecionar_arquivo(self, path: str):
        ok, msg = validar_arquivo(path)
        self._file_zone.set_file(path, ok)
        if ok:
            self._arquivo = path
            self._status.update(
                f"Arquivo carregado: {Path(path).name}",
                "success"
            )
        else:
            self._arquivo = None
            self._status.error(msg)

    def _selecionar_tipo(self, tipo_id: str):
        self._tipo = tipo_id
        for tid, card in self._cards.items():
            card.set_selected(tid == tipo_id)
        nome = next(t["nome"] for t in TIPOS if t["id"] == tipo_id)
        self._status.update(
            f"Tipo selecionado: {nome}  —  agora clique em Organizar Planilha",
            "t2"
        )

    def _iniciar(self):
        if not self._arquivo:
            self._status.warning("Selecione um arquivo antes de continuar.")
            return
        if not self._tipo:
            self._status.warning("Escolha o tipo de planilha para continuar.")
            return
        self._btn_run.configure(
            state="disabled",
            text="  ⏳  Processando...",
            fg_color=C["accent_soft"]
        )
        self._progress.start()
        self._status.processing("Processando planilha, aguarde...")
        threading.Thread(target=self._executar, daemon=True).start()

    def _executar(self):
        try:
            func  = next(t["func"]  for t in TIPOS if t["id"] == self._tipo)
            saida = gerar_nome_saida(self._arquivo, self._tipo)
            result = func(self._arquivo, saida)
            self.after(0, lambda: self._finalizar(saida, result))
        except Exception as e:
            self.after(0, lambda: self._erro(str(e)))

    def _finalizar(self, saida, r):
        self._progress.stop(success=True)
        self._btn_run.configure(
            state="normal",
            text="  ⚡  Organizar Planilha",
            fg_color=C["accent"]
        )
        linhas = r.get("linhas_finais", "?")
        self._status.success(
            f"✓  Concluído!  {linhas} linhas processadas  →  {Path(saida).name}"
        )
        os.startfile(Path(saida).parent)

    def _erro(self, msg):
        self._progress.stop(success=False)
        self._btn_run.configure(
            state="normal",
            text="  ⚡  Organizar Planilha",
            fg_color=C["accent"]
        )
        self._status.error(f"Erro: {msg}")


# ════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════
if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()