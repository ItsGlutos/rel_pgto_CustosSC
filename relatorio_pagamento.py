import xlsxwriter.utility
import re
import ctypes
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd

try:
    import customtkinter as ctk
except ImportError:
    messagebox.showerror("Erro de Dependência", "A biblioteca 'customtkinter' não está instalada.\nPor favor rode:\npip install customtkinter")
    raise SystemExit

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COMP1 = [None, None, 'St', 'Empresa', 'Fornecedo', None, 'LNeg', 'Referência', 'Nº documento', 'Data doc.', 'Doc.compensação', '   Mont.em MI', 'Vencimento', 'Texto', 'Tipo', 'GrpPrevTes', 'Compensaç.', 'BlP', 'Usuário']
COMP2 = [None, 'Empresa', 'Loc.Neg', 'Fornecedor', 'Doc.Cont', None, None, 'Dat.Doc', 'Referência', 'Dt.Pagto', 'Tp.Doc', 'BIP', 'Montant.MI', 'Texto', 'Doc.Compen', 'Dat.Compen', 'Doc.MIRO', None, 'Dt.Process', 'Msg.LOG']
COMP2_ALT = [None, 'Empresa', 'Loc.Neg', 'Fornecedor', 'Doc.Cont', None, None, 'Dat.Doc', 'Referência', 'Dt.Pagto', 'Tp.Doc', 'BIP', 'Montant.MI', 'Texto', 'Doc.Compen', 'Dat.Compen', 'Doc.Cont.MIRO', None, None, 'Dt.Process', 'Msg.LOG']
COLUNAS_PADRAO = ['Empresa', 'Fornecedor', 'LNeg', 'Referência', 'Nº documento', 'Data doc.', 'Doc.compensação', 'Valor', 'Vencimento', 'Texto', 'Tipo', 'Compensaç.', 'Blp']
TIPOS_VALIDOS = ['RE', 'KT', 'RF']

CSV_KWARGS = dict(sep='\t', header=None, encoding='utf-16')


def _formatar_df(df: pd.DataFrame, inverter_valor: bool) -> pd.DataFrame:
    for col in ('Data doc.', 'Vencimento'):
        df[col] = df[col].astype(str).str.replace('.', '/', regex=False)
    df['Valor'] = pd.to_numeric(
        df['Valor'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),
        errors='coerce'
    )
    df = df[df['Tipo'].isin(TIPOS_VALIDOS)]
    if inverter_valor:
        df['Valor'] *= -1
    return df


def tratamento(df: pd.DataFrame | None) -> pd.DataFrame | None:
    if df is None or df.empty:
        return None
    first_line = list(df.iloc[0].where(pd.notnull(df.iloc[0]), None))

    if first_line == COMP1:
        df = df.drop(df.columns[[0, 1, 2, 5, 15, 18]], axis=1).iloc[1:].reset_index(drop=True)
        df.columns = COLUNAS_PADRAO
        return _formatar_df(df, inverter_valor=True)

    if first_line in (COMP2, COMP2_ALT):
        colunas_remover = df.columns[[0, 5, 6, 16, 17, 18, 19, 20] if first_line == COMP2_ALT else [0, 5, 6, 16, 17, 18, 19]]
        df = df.iloc[1:].reset_index(drop=True).rename(columns={
            1: 'Empresa', 2: 'LNeg', 3: 'Fornecedor', 4: 'Nº documento',
            7: 'Data doc.', 8: 'Referência', 9: 'Vencimento', 10: 'Tipo',
            11: 'Blp', 12: 'Valor', 13: 'Texto', 14: 'Doc.compensação', 15: 'Compensaç.'
        }).drop(columns=colunas_remover)[COLUNAS_PADRAO]
        return _formatar_df(df, inverter_valor=False)

    return None


def _ler_linha(arq: Path, skiprows: int) -> list:
    df = pd.read_csv(arq, nrows=1, skiprows=skiprows, **CSV_KWARGS)
    return list(df.iloc[0].where(pd.notnull(df.iloc[0]), None))


def _nome_aba(valor) -> str:
    nome = re.sub(r'[\\*?:/\[\]]', '-', str(valor))[:31]
    return nome if nome and nome.lower() != 'nan' else 'Desconhecido'


def _ler_arquivo(arq: Path) -> pd.DataFrame | None:
    try:
        try:
            if _ler_linha(arq, skiprows=8) == COMP1:
                return pd.read_csv(arq, skiprows=8, dtype={8: str, 10: str}, **CSV_KWARGS)
        except Exception:
            pass
        try:
            linha = _ler_linha(arq, skiprows=3)
            if linha == COMP2:
                return pd.read_csv(arq, skiprows=3, index_col=False, names=list(range(20)), engine='python', **CSV_KWARGS)
            elif linha == COMP2_ALT:
                return pd.read_csv(arq, skiprows=3, index_col=False, names=list(range(21)), engine='python', **CSV_KWARGS)
        except Exception:
            pass
    except Exception:
        pass
    return None


def _escrever_aba(workbook, df: pd.DataFrame, sheet_name: str):
    """Escreve aba com tema verde dessaturado: cabeçalho em #4A7C59, linhas em verde-acinzentado alternado e subtotal em #2D5240."""
    VERDE, VERDE_ESC, VERDE1, VERDE2 = '#549E39', '#549E39', '#B7DFA8', '#DBEFD3'
    TEXTO = '#1A2E22'  # verde escuro dessaturado para texto nas linhas claras

    def _fmt(**kw):
        return workbook.add_format({'border': 1, 'border_color': 'white', 'valign': 'vcenter', **kw})

    fmt_cab     = _fmt(bold=True, font_color='white', bg_color=VERDE,     align='center')
    fmt_r1      = _fmt(font_color=TEXTO,              bg_color=VERDE1)
    fmt_r2      = _fmt(font_color=TEXTO,              bg_color=VERDE2)
    fmt_r1_val  = _fmt(font_color=TEXTO,              bg_color=VERDE1,    num_format='#,##0.00')
    fmt_r2_val  = _fmt(font_color=TEXTO,              bg_color=VERDE2,    num_format='#,##0.00')
    fmt_r1_val_kt = _fmt(font_color='red',            bg_color=VERDE1,    num_format='#,##0.00')
    fmt_r2_val_kt = _fmt(font_color='red',            bg_color=VERDE2,    num_format='#,##0.00')
    fmt_tot_lbl = _fmt(bold=True, font_color='white', bg_color=VERDE_ESC)
    fmt_tot_val = _fmt(bold=True, font_color='white', bg_color=VERDE_ESC, num_format='"R$ "#,##0.00')

    ws   = workbook.add_worksheet(sheet_name)
    cols = list(df.columns)
    vidx = cols.index('Valor')
    tipo_idx = cols.index('Tipo') if 'Tipo' in cols else -1
    data = df.values

    # Cabeçalho + largura automática
    for c, name in enumerate(cols):
        ws.write(0, c, name, fmt_cab)
        ws.set_column(c, c, max(len(str(name)) + 2, 14))

    # Colunas que devem ser sempre tratadas como texto (não numéricas)
    cols_texto = {c for c, name in enumerate(cols) if name.lower() in ('blp', 'tipo', 'texto', 'referência', 'nº documento', 'doc.compensação', 'compensaç.', 'empresa', 'fornecedor', 'lneg', 'data doc.', 'vencimento')}

    # Dados com linhas alternadas
    for r, row in enumerate(data, start=1):
        par = r % 2 == 0
        is_kt = tipo_idx >= 0 and row[tipo_idx] == 'KT'
        for c, val in enumerate(row):
            if c == vidx:
                fmt = (fmt_r2_val_kt if par else fmt_r1_val_kt) if is_kt else (fmt_r2_val if par else fmt_r1_val)
            else:
                fmt = fmt_r2 if par else fmt_r1
            if c in cols_texto:
                ws.write_string(r, c, '' if pd.isna(val) else str(val), fmt)
            else:
                ws.write(r, c, val, fmt)

    # Linha de subtotal
    n        = len(data)
    vcol_ref = xlsxwriter.utility.xl_col_to_name(vidx)
    for c in range(len(cols)):
        if c == 0:
            ws.write(n + 1, c, 'SUBTOTAL', fmt_tot_lbl)
        elif c == vidx:
            ws.write_formula(n + 1, c, f'=SUBTOTAL(9,{vcol_ref}2:{vcol_ref}{n + 1})', fmt_tot_val)
        else:
            ws.write(n + 1, c, '', fmt_tot_lbl)

    ws.autofilter(0, 0, n, len(cols) - 1)
    ws.freeze_panes(1, 0)


# ─── Interface Gráfica ───────────────────────────────────────────────────────

ROXO, ROXO_HOVER = "#8B5CF6", "#7C3AED"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Relatórios SAP")
        self.geometry("400x610")
        self.resizable(False, False)
        self.pasta_selecionada: Path | None = None
        self.pasta_destino: Path | None = None

        ctk.CTkLabel(self, text="Relatório de Pagamento", font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(20, 20))

        # Passo 1 — Nome do arquivo
        ctk.CTkLabel(self, text="1. Nome do arquivo a ser gerado:").pack(anchor="w", padx=30)
        self.entrada_nome = ctk.CTkEntry(self, width=340, placeholder_text="Ex: Relatorio_Jan")
        self.entrada_nome.pack(padx=30, pady=(0, 20))

        # Passo 2 — Pasta de origem
        ctk.CTkLabel(self, text="2. Selecione a pasta com os dados (.xls):").pack(anchor="w", padx=30)
        ctk.CTkButton(self, text="Procurar Pasta...", fg_color=ROXO, hover_color=ROXO_HOVER,
                      command=self.selecionar_pasta).pack(anchor="w", padx=30, pady=(5, 5))
        self.lbl_pasta_path = ctk.CTkLabel(self, text="Nenhuma pasta selecionada",
                                           text_color="gray", font=ctk.CTkFont(size=11, slant="italic"))
        self.lbl_pasta_path.pack(anchor="w", padx=30, pady=(0, 20))

        # Passo 3 — Separador de abas
        ctk.CTkLabel(self, text="3. Dividir as abas do Excel por:").pack(anchor="w", padx=30)
        self.sep_var = ctk.StringVar(value="Empresa")
        for opcao in ("Empresa", "Fornecedor", "LNeg"):
            ctk.CTkRadioButton(self, text=opcao, variable=self.sep_var, value=opcao,
                               fg_color=ROXO, hover_color=ROXO_HOVER).pack(anchor="w", padx=30, pady=5)

        # Passo 4 — Pasta de destino
        ctk.CTkLabel(self, text="4. Selecione a pasta de destino do relatório:").pack(anchor="w", padx=30, pady=(20, 0))
        ctk.CTkButton(self, text="Procurar Pasta...", fg_color=ROXO, hover_color=ROXO_HOVER,
                      command=self.selecionar_pasta_destino).pack(anchor="w", padx=30, pady=(5, 5))
        self.lbl_destino_path = ctk.CTkLabel(
            self,
            text="Nenhuma pasta de destino selecionada",
            text_color="gray",
            font=ctk.CTkFont(size=11, slant="italic")
        )
        self.lbl_destino_path.pack(anchor="w", padx=30, pady=(0, 0))

        # Botão processar
        self.btn_processar = ctk.CTkButton(
            self, text="PROCESSAR E GERAR", fg_color=ROXO, hover_color=ROXO_HOVER,
            font=ctk.CTkFont(weight="bold", size=14), command=self.iniciar_processamento
        )
        self.btn_processar.pack(fill="x", padx=30, pady=(25, 0), ipady=8)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Qual pasta deseja usar?")
        if pasta:
            self.pasta_selecionada = Path(pasta)
            display = str(self.pasta_selecionada)
            self.lbl_pasta_path.configure(
                text=("..." + display[-37:] if len(display) > 40 else display),
                text_color="white"
            )

    def selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(title="Onde deseja salvar o relatório?")
        if pasta:
            self.pasta_destino = Path(pasta)
            display = str(self.pasta_destino)
            self.lbl_destino_path.configure(
                text=("..." + display[-37:] if len(display) > 40 else display),
                text_color="white"
            )

    def _resetar_botao(self):
        self.btn_processar.configure(state="normal", text="PROCESSAR E GERAR")

    def iniciar_processamento(self):
        if not self.entrada_nome.get():
            messagebox.showwarning("Aviso", "Por favor, defina o nome do arquivo final no passo 1.")
            return
        if not self.pasta_selecionada:
            messagebox.showwarning("Aviso", "Por favor, selecione a pasta de origem dos arquivos .xls no passo 2.")
            return
        self.btn_processar.configure(state="disabled", text="Processando Arquivos... Aguarde")
        threading.Thread(
            target=self.processar_dados,
            args=(self.entrada_nome.get(), self.pasta_selecionada, self.sep_var.get()),
            daemon=True
        ).start()

    def processar_dados(self, nome_str: str, pasta_path: Path, agrupador: str):
        # Usa a pasta de destino escolhida, ou o Desktop como fallback
        destino = self.pasta_destino if self.pasta_destino else print("Erro")
        try:
            arquivo_final = destino / Path(nome_str).with_suffix('.xlsx')
            arqvs = list(pasta_path.rglob('*.xls'))

            if not arqvs:
                messagebox.showerror('Erro', f'Nenhum arquivo .xls encontrado em:\n{pasta_path}')
                self._resetar_botao()
                return

            lista_dfs = [df for arq in arqvs if (df := tratamento(_ler_arquivo(arq))) is not None]  # type: ignore[misc]

            if not lista_dfs:
                messagebox.showwarning('Aviso', 'Nenhum dado válido encontrado após o processamento.')
                self._resetar_botao()
                return

            df_fim = pd.concat(lista_dfs, ignore_index=True)

            workbook = xlsxwriter.Workbook(str(arquivo_final), {'nan_inf_to_errors': True})
            for valor in df_fim[agrupador].unique():
                _escrever_aba(workbook, df_fim[df_fim[agrupador] == valor].reset_index(drop=True), _nome_aba(valor))
            workbook.close()

            self.after(0, lambda: messagebox.showinfo('Processo Finalizado!', f'{arquivo_final.stem} foi gerado com sucesso!'))
            self.after(0, self._resetar_botao)

        except Exception as e:
            self.after(0, lambda: messagebox.showinfo('Processo Finalizado!', f'{arquivo_final.stem} foi gerado com sucesso!'))
            self.after(0, self._resetar_botao)


if __name__ == "__main__":
    App().mainloop()