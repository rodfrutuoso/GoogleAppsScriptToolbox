# Importador de Relatório de Chamados do Desk

## Descrição

Este script foi criado para importar o Relatório de Chamados do Help Desk por meio de um link URL da integração do Power BI. Ele utiliza a classe UrlFetchApp para consultar o HTML da página e extrair os dados relevantes.

## Funcionalidades

O script possui as seguintes funcionalidades:

- **Importação do Relatório**: O script acessa o relatório de chamados do Desk por meio do link do Power BI e extrai os dados em formato tabular.

- **Verificação de Novos Chamados**: Ele verifica quais chamados ainda não estão na planilha do Google Sheets e os insere automaticamente.

- **Ajuste de Limite de Caracteres**: O script ajusta o tamanho das strings para garantir que elas não excedam o limite de 5000 caracteres por célula no Google Sheets.

## Como Usar

1. Certifique-se de que os requisitos sejam atendidos (veja a seção de Requisitos).

2. Atualize as variáveis `url`, `abaRange` e `abaRangeAux` de acordo com sua planilha e url do Relatório do Help Desk.
   
3.  No Google Sheets, execute a função `filtroChamadosNovos` para verificar e inserir novos chamados.

## Requisitos

Antes de usar este script, verifique se os seguintes requisitos estão satisfeitos:

### 1. Acesso ao Relatório do Desk

Certifique-se de que você tenha acesso ao Relatório de Chamados do Help Desk por meio do link do Power BI especificado no código.

### 2. API do Google Sheets

Instalar a API do Google Sheets através do menu de "Serviços" do Google Apps Script.

## Autor

Este script foi desenvolvido por [Rodrigo Frutuoso](https://rodfrutuoso.github.io/teste/) .

