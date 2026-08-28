import customtkinter as ctk
from tkinter import filedialog
import os
import threading
from tkinter import messagebox

# [gerado] Importações lazy serão feitas diretamente dentro das funções que as usam
# para acelerar a abertura da interface gráfica

# ── Tema e aparência ──────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Paleta WCAG AA (contraste ≥ 4.5:1 sobre fundos escuros) ──────────────────
NAVY_BG       = "#0D1B2A"   # fundo principal
NAVY_SURFACE  = "#112236"   # superfície dos cards
NAVY_BORDER   = "#1B3A5C"   # bordas / divisores
ACCENT_BLUE   = "#2D7DD2"   # azul destaque (botões primários)
ACCENT_HOVER  = "#3A8FE8"   # hover dos botões primários
TEXT_PRIMARY  = "#E8F0FE"   # texto principal  (contraste ~12:1)
TEXT_SECONDARY= "#94B4D4"   # texto secundário (contraste ~5:1)
FIELD_BG      = "#0A1828"   # fundo dos inputs
SUCCESS_GREEN = "#2ECC71"   # feedback de sucesso (não usado aqui, reservado)
RADIUS        = 10


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ── Janela ────────────────────────────────────────────────────────────
        self.title("PagFlow Excel")
        self.resizable(False, False)
        self.configure(fg_color=NAVY_BG)

        # Centralizar na tela
        self.update_idletasks()
        w, h = 500, 610  # [compacto] Reduzido para caber em telas menores
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        self._build_ui()

    # ── Construção da UI ──────────────────────────────────────────────────────
    def _build_ui(self):
        # Padding externo — [compacto] reduzido de 28/24 para 20/14
        outer = ctk.CTkFrame(self, fg_color=NAVY_BG)
        outer.pack(fill="both", expand=True, padx=20, pady=14)

        # ── Cabeçalho ─────────────────────────────────────────────────────────
        header = ctk.CTkFrame(outer, fg_color=NAVY_SURFACE,
                              corner_radius=RADIUS, border_width=1,
                              border_color=NAVY_BORDER, height=38)  # [compacto] 45 → 38
        header.pack(fill="x", pady=(0, 12))  # [compacto] 18 → 12
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="PagFlow",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),  # [compacto] 18 → 16
            text_color=TEXT_PRIMARY,
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            header,
            text="Feito por Gustavo Pinheiro",
            font=ctk.CTkFont(family="Segoe UI", size=10),  # [compacto] 11 → 10
            text_color=TEXT_SECONDARY,
        ).pack(side="right", padx=12, pady=10)  # [compacto] padx 16 → 12

        # ── Campo 1 — Nome do arquivo ─────────────────────────────────────────
        self._section_label(outer, "1", "Nome do arquivo de saída")

        self.entry_nome = self._styled_entry(
            outer, placeholder="ex.: relatorio_outubro"
        )

        # ── Campo 2 — Pasta de origem ─────────────────────────────────────────
        self._section_label(outer, "2", "Pasta com os arquivos de entrada")

        row2 = ctk.CTkFrame(outer, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 10))  # [compacto] 14 → 10

        self.entry_pasta = self._styled_entry(
            row2, placeholder="Selecione a pasta...", pack=False
        )
        self.entry_pasta.pack(side="left", fill="x", expand=True)

        self._browse_btn(row2, self._browse_pasta).pack(side="left", padx=(6, 0))

        # ── Campo 3 — Separação ───────────────────────────────────────────────
        self._section_label(outer, "3", "Separação das abas")

        seg_frame = ctk.CTkFrame(outer, fg_color="transparent")
        seg_frame.pack(fill="x", pady=(0, 10))  # [compacto] 14 → 10

        self.sep_var = ctk.StringVar(value="Não Separar")
        opcoes = ["Marca", "LNeg", "Fornecedor", "Não Separar"]
        self.seg = ctk.CTkSegmentedButton(
            seg_frame,
            values=opcoes,
            variable=self.sep_var,
            fg_color=NAVY_SURFACE,
            selected_color=ACCENT_BLUE,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=NAVY_SURFACE,
            unselected_hover_color=NAVY_BORDER,
            text_color=TEXT_PRIMARY,
            text_color_disabled=TEXT_SECONDARY,
            font=ctk.CTkFont(family="Segoe UI", size=11),  # [compacto] 12 → 11
            corner_radius=8,
            border_width=1,
        )
        self.seg.pack(fill="x")

        # ── Campo 4 — Personalização (cor da tabela) ──────────────────────────
        self._section_label(outer, "4", "Personalização")

        self.cor_var = ctk.StringVar(value="Azul")
        opcoes_cor = ["Verde", "Azul", "Roxo", "Vermelho"]
        self.opt_cor = ctk.CTkOptionMenu(
            outer,
            values=opcoes_cor,
            variable=self.cor_var,
            fg_color=NAVY_SURFACE,
            button_color=ACCENT_BLUE,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=NAVY_SURFACE,
            dropdown_hover_color=NAVY_BORDER,
            dropdown_text_color=TEXT_PRIMARY,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family="Segoe UI", size=11),  # [compacto] 12 → 11
            corner_radius=RADIUS,
            height=32,  # [compacto] 38 → 32
        )
        self.opt_cor.pack(fill="x", pady=(0, 10))  # [compacto] 14 → 10

        # ── Campo 5 — Pasta de destino ────────────────────────────────────────
        self._section_label(outer, "5", "Destino do arquivo final")

        row4 = ctk.CTkFrame(outer, fg_color="transparent")
        row4.pack(fill="x", pady=(0, 12))  # [compacto] 18 → 12

        self.entry_destino = self._styled_entry(
            row4, placeholder="Selecione onde salvar...", pack=False
        )
        self.entry_destino.pack(side="left", fill="x", expand=True)

        self._browse_btn(row4, self._browse_destino).pack(side="left", padx=(6, 0))

        # ── Container Inferior ────────────────────────────────────────────────
        bottom_container = ctk.CTkFrame(outer, fg_color="transparent")
        bottom_container.pack(side="bottom", fill="x")

        # ── Divisor ───────────────────────────────────────────────────────────
        div = ctk.CTkFrame(bottom_container, fg_color=NAVY_BORDER, height=1)
        div.pack(fill="x", pady=(2, 12))  # [compacto] 4/16 → 2/12

        # ── Botões de ação ────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(bottom_container, fg_color="transparent")
        btn_row.pack(fill="x")

        ctk.CTkButton(
            btn_row,
            text="Limpar campos",
            width=120,  # [compacto] 130 → 120
            height=34,  # [compacto] 40 → 34
            fg_color=NAVY_SURFACE,
            hover_color=NAVY_BORDER,
            text_color=TEXT_SECONDARY,
            border_width=1,
            border_color=NAVY_BORDER,
            corner_radius=RADIUS,
            font=ctk.CTkFont(family="Segoe UI", size=12),  # [compacto] 13 → 12
            command=self._limpar,
        ).pack(side="left")

        ctk.CTkButton(
            btn_row,
            text="▶  Gerar Excel",
            height=34,  # [compacto] 40 → 34
            fg_color=ACCENT_BLUE,
            hover_color=ACCENT_HOVER,
            text_color="#FFFFFF",
            corner_radius=RADIUS,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),  # [compacto] 13 → 12
            command=self._gerar,
        ).pack(side="right")

        # ── Bloco de status ───────────────────────────────────────────────────
        self.status_frame = ctk.CTkFrame(
            bottom_container,
            fg_color="#1B3A5C",
            corner_radius=8,
            border_width=0,
            height=30,  # [compacto] 36 → 30
        )
        self.status_frame.pack(fill="x", pady=(10, 0))  # [compacto] 14 → 10
        self.status_frame.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(
            self.status_frame,
            text="Pronto para uso.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_SECONDARY,
            anchor="center",
        )
        self.lbl_status.pack(fill="x", padx=12, pady=6)  # [compacto] pady 8 → 6

    # ── Helpers de UI ─────────────────────────────────────────────────────────
    def _section_label(self, parent, num: str, text: str):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 3))  # [compacto] 4 → 3

        ctk.CTkLabel(
            row,
            text=num,
            width=20, height=20,  # [compacto] 22/22 → 20/20
            fg_color=ACCENT_BLUE,
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),  # [compacto] 10 → 9
            corner_radius=10,  # [compacto] 11 → 10
        ).pack(side="left")

        ctk.CTkLabel(
            row,
            text=f"  {text}",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),  # [compacto] 12 → 11
            text_color=TEXT_PRIMARY,
        ).pack(side="left")

    def _styled_entry(self, parent, placeholder: str, pack: bool = True):
        entry = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            placeholder_text_color=TEXT_SECONDARY,
            fg_color=FIELD_BG,
            border_color=NAVY_BORDER,
            text_color=TEXT_PRIMARY,
            border_width=1,
            corner_radius=RADIUS,
            height=32,  # [compacto] 38 → 32
            font=ctk.CTkFont(family="Segoe UI", size=11),  # [compacto] 12 → 11
        )
        if pack:
            entry.pack(fill="x", pady=(0, 10))  # [compacto] 14 → 10
        return entry

    def _browse_btn(self, parent, cmd):
        return ctk.CTkButton(
            parent,
            text="📁",
            width=36, height=32,  # [compacto] 40/38 → 36/32
            fg_color=NAVY_SURFACE,
            hover_color=NAVY_BORDER,
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=NAVY_BORDER,
            corner_radius=RADIUS,
            font=ctk.CTkFont(size=13),  # [compacto] 14 → 13
            command=cmd,
        )

    # ── Callbacks ─────────────────────────────────────────────────────────────
    def _browse_pasta(self):
        path = filedialog.askdirectory(title="Selecionar pasta de entrada")
        if path:
            self.entry_pasta.delete(0, "end")
            self.entry_pasta.insert(0, path)
            self._set_status(f"Pasta de entrada: {os.path.basename(path)}")

    def _browse_destino(self):
        path = filedialog.askdirectory(title="Selecionar pasta de destino")
        if path:
            self.entry_destino.delete(0, "end")
            self.entry_destino.insert(0, path)
            self._set_status(f"Destino definido: {os.path.basename(path)}")

    def _limpar(self):
        for entry in (self.entry_nome, self.entry_pasta, self.entry_destino):
            entry.delete(0, "end")
        self.sep_var.set("Não Separar")
        self.cor_var.set("Azul")
        self._set_status("Campos limpos.")

    def _gerar(self):
        nome     = self.entry_nome.get().strip()
        pasta    = self.entry_pasta.get().strip()
        separar  = self.sep_var.get()
        cor      = self.cor_var.get()
        destino  = self.entry_destino.get().strip()

        if not nome or not pasta or not destino:
            campos = []
            if not nome:    campos.append('Nome do arquivo')
            if not pasta:   campos.append('Pasta de entrada')
            if not destino: campos.append('Destino')
            self._set_status(f"⚠  Campo(s) não preenchido(s): {', '.join(campos)}", estado='aviso')
            return

        def _rodar():
            self._set_status(f'⏳  Processando... · Separação: {separar}')

            try:
                from ope import executar
                executar(
                    pasta=pasta,
                    destino=destino,
                    nome=nome,
                    separar=separar,
                    cor=cor,
                    cb_status=self._set_status,
                )
            except Exception as e:
                self._set_status(f'✗  Erro fatal: {e}', estado='erro')

        threading.Thread(target=_rodar, daemon=True).start()

    def _set_status(self, msg: str, estado: str = None, warn: bool = False):
        if warn and estado is None:
            estado = 'aviso'

        config = {
            'ok':    ('#1A6B35', '#FFFFFF'),
            'erro':  ('#8B1A1A', '#FFFFFF'),
            'aviso': ('#B8860B', '#FFFFFF'),
        }
        bg, fg = config.get(estado, ('#1B3A5C', TEXT_SECONDARY))

        def _aplicar(b=bg, f=fg, m=msg, est=estado):
            if est == 'ok':
                from tkinter import messagebox
                messagebox.showinfo("Sucesso", m)
                self.status_frame.configure(fg_color='#1B3A5C')
                self.lbl_status.configure(text="Pronto para uso.", text_color=TEXT_SECONDARY)
            else:
                self.status_frame.configure(fg_color=b)
                self.lbl_status.configure(text=m, text_color=f)

        self.after(0, _aplicar)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()