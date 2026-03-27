from pathlib import Path
from tkinter import Tk, filedialog, messagebox, simpledialog
import os
import ctypes
import openpyxl

ctypes.windll.shcore.SetProcessDpiAwareness(1)

root = Tk()
root.withdraw()
root.attributes('-topmost', True)

# comparação entre a primeira linha de cada relatório:
# fbl1n
comp1 = [None, None, 'St', 'Empresa', 'Fornecedo', None, 'LNeg', 'Referência', 'Nº documento', 'Data doc.', 'Doc.compensação', '   Mont.em MI', 'Vencimento', 'Texto', 'Tipo', 'GrpPrevTes', 'Compensaç.', 'BlP', 'Usuário']
# rel_log
comp2 = [None, 'Empresa', 'Loc.Neg', 'Fornecedor', 'Doc.Cont', None, None, 'Dat.Doc', 'Referência', 'Dt.Pagto', 'Tp.Doc', 'BIP', 'Montant.MI', 'Texto', 'Doc.Compen', 'Dat.Compen', 'Doc.MIRO', None, 'Dt.Process', 'Msg.LOG']

possible_desktops = list(Path.home().glob("**/Desktop"))

nome_arquivo = Path(simpledialog.askstring('Nome do Arquivo','Escolha o nome do seu arquivo')).with_suffix('.xlsx')

pasta = Path(filedialog.askdirectory(title='Qual pasta deseja usar?'))

#para onde o arquivo será salvo
salvo_em = possible_desktops

arquivo_final = salvo_em[0] / nome_arquivo
arqvs = list(pasta.rglob('*.xls'))

def tratamento(df):
    first_line = list(df.iloc[0].where(pd.notnull(df.iloc[0]), None))
    colunas_padrao = [
        'Empresa', 'Fornecedor', 'LNeg', 'Referência', 'Nº documento', 
        'Data doc.', 'Doc.compensação', 'Valor', 'Vencimento', 'Texto', 
        'Tipo', 'Compensaç.', 'Blp'
    ]

    
    if first_line == comp1:

        df = df.drop(df.columns[[0, 1, 2, 5, 15, 18]], axis=1)
        df = df.drop(index=0).reset_index(drop=True)

        # define o cabeçalho
        df.columns = colunas_padrao

        df['Data doc.'] = df['Data doc.'].str.replace('.', '/')
        df['Vencimento'] = df['Vencimento'].str.replace('.','/')

        df['Valor'] = df['Valor'].str.replace('.','')
        df['Valor'] = df['Valor'].str.replace(',','.')
        df['Valor'] = df['Valor'].astype(float)

        df = df[df['Tipo'].isin(['RE', 'KT','RF'])]
        df['Valor'] = df['Valor']*(-1)

    elif comp2 == first_line:

        colunas_para_remover = df.columns[[0, 5, 6, 16, 17, 18, 19]]
        df = df.drop(index=0).reset_index(drop=True)

        colunas_renomeadas = {2:'LNeg',
        9:'Vencimento',
        10:'Tipo',
        11:'Blp',
        15:'Compensaç.',
        12:'Valor',
        14:'Doc.compensação',
        7:'Data doc.',
        3:'Fornecedor',
        4:'Nº documento',
        13:'Texto',
        8:'Referência',
        1:'Empresa'}

        df = df.rename(columns=colunas_renomeadas)

        df = df.drop(columns=colunas_para_remover)

        df = df[['Empresa', 'Fornecedor', 'LNeg', 'Referência', 'Nº documento', 'Data doc.', 'Doc.compensação', 'Valor', 'Vencimento', 'Texto', 'Tipo', 'Compensaç.','Blp']]

        df['Data doc.'] = df['Data doc.'].str.replace('.', '/')
        df['Vencimento'] = df['Vencimento'].str.replace('.','/')

        df['Valor'] = df['Valor'].str.replace('.','')
        df['Valor'] = df['Valor'].str.replace(',','.')
        df['Valor'] = df['Valor'].astype(float)

        df = df[df['Tipo'].isin(['RE', 'KT','RF'])]

    else:
        messagebox.showerror("ERRO", "FORMATO DE ARQUIVO NÃO RECONHECIDO")
        return None

    return df

import pandas as pd

try:

    lista_dfs = []

    qtd_colunas = list(range(20))

    for arq in arqvs:

        df_probe1 = pd.read_csv(arq, sep='\t', header=None, encoding='utf-16', nrows=1, skiprows=8)
        df_probe2 = pd.read_csv(arq, sep='\t', header=None, encoding='utf-16', nrows=1, skiprows=3)

        first_line1 = list(df_probe1.iloc[0].where(pd.notnull(df_probe1.iloc[0]), None))
        first_line2 = list(df_probe2.iloc[0].where(pd.notnull(df_probe2.iloc[0]), None))

        if first_line1 == comp1:

            df = pd.read_csv(arq, sep='\t',header=None ,encoding='utf-16', skiprows=8, dtype={
                    8:str,
                    10:str
                })
        elif first_line2 == comp2:

            df = pd.read_csv(arq, sep='\t', skiprows=3, encoding='utf-16', index_col=False, header=None, names=qtd_colunas, engine='python')
            # df = df.drop(index=0).reset_index(drop=True)
            
        else:
            # messagebox.showerror('Erro Na consolidação',f'Formato de arquivo não reconhecido:\n{arq.name}')
            continue

        df = tratamento(df)
        lista_dfs.append(df)

    df_fim = pd.concat(lista_dfs, ignore_index=True)

    with pd.ExcelWriter(arquivo_final, engine='openpyxl') as writer:
        
        for marca in df_fim['Empresa'].unique():
            df_filtrado = df_fim[df_fim['Empresa'] == marca]
            df_filtrado.to_excel(writer, sheet_name = marca, index=False)

    import xlwings as xw

    arq_xlsx = str(arquivo_final)
    arq_xlsb = arq_xlsx.replace(".xlsx", ".xlsb")

    app = xw.App(visible=False)
    wb = app.books.open(arquivo_final)
    wb.save(arq_xlsb)
    wb.close()
    app.quit()

    os.remove(arq_xlsx)
    messagebox.showinfo('Relatório de Pagamento', f'{nome_arquivo} foi gerado com sucesso')

except Exception as e:
    messagebox.showerror('Erro', f'Seu arquivo não pôde ser gerado:\n{e}')