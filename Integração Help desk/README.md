# Importador de Relatório de Chamados do Desk

# Desk Ticket Report Importer

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


## Description

This script was created to import the Desk Ticket Report through a Power BI integration URL link. It uses the UrlFetchApp class to query the HTML of the page and extract the relevant data.

## Features

* **Report Import:** The script accesses the Desk ticket report through the Power BI link and extracts the data in tabular format.
* **New Ticket Check:** It checks which tickets are not yet in the Google Sheets spreadsheet and inserts them automatically.
* **Character Limit Adjustment:** The script adjusts the size of strings to ensure they do not exceed the 5000 character limit per cell in Google Sheets.

## How to Use

1. Make sure that the requirements are met (see the Requirements section).
2. Update the `url`, `abaRange`, and `abaRangeAux` variables according to your spreadsheet and Desk Ticket Report URL.
3. In Google Sheets, run the `filtroChamadosNovos` function to check and insert new tickets.

## Requirements

Before using this script, make sure that the following requirements are met:

### 1. Desk Report Access

Make sure that you have access to the Desk Ticket Report through the Power BI link specified in the code.
### 2. Google Sheets API

Install the Google Sheets API through the "Services" menu in Google Apps Script.

## Author

This script was developed by [Rodrigo Frutuoso](https://rodfrutuoso.github.io/teste/).

