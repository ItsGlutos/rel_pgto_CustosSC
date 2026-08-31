# PagFlow

Aplicação desktop desenvolvida para automatizar a geração de relatórios de pagamento de fretes, com base em Conhecimentos de Transporte Eletrônico (CT-e) e Notas Fiscais.

## Sobre o projeto

O PagFlow processa relatórios extraídos do SAP (transações FBL1N e ZFI0678_REL_LOG) e consolida as informações em uma planilha Excel formatada, destinada à conferência e ao fechamento do pagamento de fretes. A aplicação conta com interface gráfica desenvolvida em CustomTkinter, priorizando simplicidade operacional para usuários sem conhecimento técnico.

## Funcionalidades

- Consolidação de relatórios extraídos do SAP (FBL1N e ZFI0678_REL_LOG)
- Deduplicação de lançamentos de CT-e
- Separação opcional das abas do relatório por Marca, LNeg ou Fornecedor
- Personalização visual da planilha gerada
- Processamento em segundo plano, sem interrupção da interface durante a geração
- Retorno de status em tempo real (sucesso, erro, aviso)

## Estrutura do projeto

O projeto é organizado em três módulos principais.

`main.py` contém a interface gráfica da aplicação, responsável pela coleta dos parâmetros informados pelo usuário: nome do arquivo de saída, pasta de origem, critério de separação, personalização visual e pasta de destino.

`ope.py` é responsável pela orquestração do fluxo de processamento, acionado a partir da interface.

`funcoes.py` concentra a lógica de negócio da aplicação, incluindo o tratamento dos dados extraídos do SAP.

## Requisitos

- Python 3.x
- CustomTkinter
- Pandas
- Openpyxl
- Layout FBL1N: /RELPGTCUSTO

## Como usar

Execute a aplicação (ou o executável gerado), informe o nome do arquivo de saída, selecione a pasta com os relatórios extraídos do SAP, defina o critério de separação e a personalização visual desejados, selecione a pasta de destino e clique em **Gerar Excel**.

## Distribuição

A aplicação é empacotada com PyInstaller e distribuída por meio de instalador, viabilizando o uso em ambientes sem configuração prévia de Python.

## Autor

Desenvolvido por Gustavo Pinheiro.
