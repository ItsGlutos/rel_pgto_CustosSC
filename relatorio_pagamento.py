from pathlib import Path
from tkinter import Tk, filedialog, messagebox, simpledialog
import os
import ctypes
import openpyxl

ctypes.windll.shcore.SetProcessDpiAwareness(1)

root = Tk()
root.withdraw()
root.attributes('-topmost', True)

possible_desktops = list(Path.home().glob("**/Desktop"))

nome_arquivo = Path(simpledialog.askstring('Nome do Arquivo','Escolha o nome do seu arquivo')).with_suffix('.xlsx')

# pasta = Path(filedialog.askdirectory(title='Qual pasta deseja usar?'))
pasta = Path(r"C:\\Users\\gustavo.pinheiro\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\Automações pessoais\\Relatórios de pagamento")

#para onde o arquivo será salvo
# salvo_em = filedialog.askdirectory(title='Onde deseja salvar o seu relatório?')
# salvo_em = Path(r"C:\\Users\\gustavo.pinheiro\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\\Desktop\\")
salvo_em = possible_desktops

arquivo_final = salvo_em[0] / nome_arquivo
print(arquivo_final)

arqvs = list(pasta.rglob('*.xls'))

def tratamento(df):
    
    df = df.drop(df.columns[[0, 1, 2, 5]], axis=1)

    # define o cabeçalho
    df.columns = ['Empresa', 'Fornecedor', 'LNeg', 'Referêcia', 'Nº documento', 'Data doc.', 'Doc.compensação', 'Valor', 'Vencimento', 'Texto', 'Tipo', 'GrpPrevTes', 'Compensaç.','Blp', 'Usuário']

    df['Data doc.'] = df['Data doc.'].str.replace('.', '/')
    df['Vencimento'] = df['Vencimento'].str.replace('.','/')

    df['Valor'] = df['Valor'].str.replace('.','')
    df['Valor'] = df['Valor'].str.replace(',','.')
    df['Valor'] = df['Valor'].astype(float)

    df = df[df['Tipo'].isin(['RE', 'KT','RF'])]
    df['Valor'] = df['Valor']*(-1)

    return df

try:
    

    lista_dfs = []

    import pandas as pd

    for arq in arqvs:
        df = pd.read_csv(arq, sep='\t',header=None ,encoding='utf-16', skiprows=9, dtype={
                8:str,
                10:str
            })
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

    