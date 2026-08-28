from datetime import datetime
from pathlib import Path


# [gerado] ─────────────────────────────────────────────────────────────────────
# FUNÇÃO: executar
# Objetivo: Encapsula todo o fluxo de automação para que possa ser chamada
#           tanto pela interface gráfica (interface.py) quanto diretamente
#           via linha de comando.

#
# Parâmetros:
#   pasta   – str ou Path com a pasta que contém os arquivos .xls de entrada
#   destino – str ou Path com a pasta onde o .xlsx final será salvo
#   nome    – nome base do arquivo de saída (sem extensão)
#   separar – opção de separação de abas vinda da interface:
#             'Marca' | 'LNeg' | 'Fornecedor' | 'Não Separar'
#   cb_status – callback opcional (função) para enviar mensagens de progresso
#               à interface; se None, usa print()
# [gerado] ─────────────────────────────────────────────────────────────────────
def executar(pasta, destino, nome, separar='Não Separar', cor='Azul', cb_status=None):

    # [gerado] Lazy imports: carrega o pandas e o funcoes.py apenas na hora da execução
    import pandas as pd
    from funcoes import processar_arquivo, exportar_excel

    # [gerado] Usa print como fallback se nenhum callback for fornecido.
    # Repassa o "estado" (ok, erro, aviso) para o main.py pintar a barra.
    def log(msg, estado=None):
        if cb_status:
            cb_status(msg, estado=estado)
        else:
            print(msg)

    pasta = Path(pasta)

    # Coleta recursivamente todos os arquivos .xls dentro da pasta
    # (rglob garante que subpastas também sejam varridas)
    arquivos = list(pasta.rglob('*.xls'))

    if not arquivos:
        log('⚠  Nenhum arquivo .xls encontrado na pasta selecionada.', estado='aviso')
        return

    # ── Loop principal ────────────────────────────────────────────────────────
    # Para cada arquivo encontrado:
    #   1. Chama processar_arquivo() (funcoes.py) que detecta o tipo (A ou B)
    #      e retorna um DataFrame já limpo e filtrado
    #   2. Adiciona o resultado na lista de tabelas tratadas
    #   3. Ao final, concatena tudo em um único df_final

    tabelas_tratadas = []   # acumula os DataFrames de cada arquivo

    for caminho in arquivos:
        try:
            # processar_arquivo detecta o relatório e aplica o tratamento adequado
            # retorna um DataFrame padronizado com as mesmas colunas em ambos os tipos
            df_tratado = processar_arquivo(caminho)

            # Só adiciona se o DataFrame não vier vazio após os filtros de tipo
            if not df_tratado.empty:
                tabelas_tratadas.append(df_tratado)
            else:
                log(f'⚠ Nenhuma linha válida (RE/KT/RF) em {caminho.name}', estado='aviso')

        except Exception as e:
            # Registra o erro na interface com cor vermelha
            log(f'✗ Erro ao processar {caminho.name}: {e}', estado='erro')

    # ── Consolida todos os DataFrames em um único resultado ───────────────────
    if tabelas_tratadas:
        # ignore_index reordena o índice de 0 a N para evitar duplicidade entre arquivos
        df_final = pd.concat(tabelas_tratadas, ignore_index=True)

        # Remove CTes duplicados
        df_final = df_final.drop_duplicates(subset=['Referência', 'Empresa', 'LNeg'])

        # [gerado] Chama exportar_excel() (funcoes.py) que cuida da separação
        # por abas conforme a opção escolhida na interface e salva o arquivo
        caminho_saida = exportar_excel(df_final, separar, destino, nome, cor)
        log(f'"{nome}" gerado com sucesso!', estado='ok')

    else:
        log('⚠  Nenhum arquivo foi processado com sucesso.', estado='aviso')


# [gerado] ─────────────────────────────────────────────────────────────────────
# Execução direta via linha de comando (sem interface gráfica)
# Mantém o comportamento anterior: usa a pasta hardcoded e salva no mesmo local
# [gerado] ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':

    pasta   = Path(r'C:\Users\gusta\OneDrive\Docs teste')
    destino = pasta                                          # salva na própria pasta
    nome    = f'consolidado_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

    executar(pasta, destino, nome, separar='Não Separar')