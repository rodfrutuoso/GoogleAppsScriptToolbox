# Listagem de arquivos de pasta do Google Drive e extração de dados de arquivo Excel

# File Listing from do Google Drive folder and extracting data from Excel file

## Descrição

Esses scripts realizam a listagem de todos os arquivos de uma pasta específica e lança esses dados em uma planilha do Google Sheets. Através das informações de tipo de arquivo e ID é feito identificação de quais arquivos estão no formato Excel e, através da biblioteca **SheetsJS**, é realizado a extração e organização dos dados dessas planilhas em uma aba.

## Funcionalidades

O conjunto de scripts inclui as seguintes funcionalidades:

- **Criação de Menu Personalizado**: Cria um menu na planilha para acionar os scripts de forma fácil e rápida.

- **Atualizar Arquivos**: Utiliza a API do Google Drive para listar todos os arquivos e pastas de uma pasta específica e coloca essa relação em uma aba chamada "RELAÇÃO ATUAL".

- **Extrair Dados de Abas em Excel**: Obtém os IDs de links de arquivos em formato Excel na aba "RELAÇÃO ATUAL" e, em seguida, abre esses arquivos, converte-os em CSV e insere os dados na aba "CONTEÚDOS" da planilha atual.

## Como Usar

1. Insira os dois scripts .js dessa pasta na planilha no Google Sheets onde deseja utilizar esses scripts.

2. Atualize as variáveis `abaExcel` `range_sheets_cola_AUX` `range_sheets_cola` `range_sheets_copia` `abaListarArquivos` `folderId` de acordo com sua necessidade

3. Acesse o menu "Scripts" criado.

4. Execute "Atualizar Arquivos" e depois "Extrair Dados" para puxar as informações

## Requisitos

Antes de utilizar esses scripts, certifique-se de que os seguintes requisitos sejam atendidos:

### 1. API do Google Sheets

Instalar a API do Google Sheets através do menu de "Serviços" do Google Apps Script

### 2. API do Google Drive

Instalar a API do Google Drive através do menu de "Serviços" do Google Apps Script

### 3. Adicionar a Biblioteca SheetsJS

Criar um arquivo do tipo .gs no Google Apps Script e renomear ele para "SheetsJS" e inserir os dados que estão no arquivo SheetsJS.js dessa pasta

## Autor

Estes scripts foram desenvolvidos por [Rodrigo Frutuoso](https://github.com/rodfrutuoso) para facilitar a automação de tarefas em planilhas do Google Sheets.

<br>
<br>

---

<br>
<br>


## Description

These scripts perform the listing of all files in a specific folder and place this data in a Google Sheets spreadsheet. They identify Excel files based on file type and ID, then use the **SheetsJS** library to extract and organize the data from these spreadsheets into a sheet.

## Features

The script bundle includes the following features:

- **Custom Menu Creation**: Creates a menu in the spreadsheet for easy and quick script execution.

- **Update Files**: Uses the Google Drive API to list all files and folders in a specific folder and places this list in a sheet called "CURRENT LIST."

- **Extract Excel Sheet Data**: Retrieves the IDs of Excel format file links in the "CURRENT LIST" sheet and then opens these files, converts them to CSV, and inserts the data into the "CONTENTS" sheet of the current spreadsheet.

## How to Use

1. Insert both .js scripts from this folder into the Google Sheets spreadsheet where you want to use these scripts.

2. Update the variables `excelSheet`, `range_sheets_cola_AUX`, `range_sheets_cola`, `range_sheets_copia`, `listFilesSheet`, and `folderId` according to your needs.

3. Access the created "Scripts" menu.

4. Execute "Update Files" and then "Extract Data" to retrieve the information.

## Requirements

Before using these scripts, make sure the following requirements are met:

### 1. Google Sheets API

Install the Google Sheets API through the "Services" menu of Google Apps Script.

### 2. Google Drive API

Install the Google Drive API through the "Services" menu of Google Apps Script.

### 3. Add the SheetsJS Library

Create a .gs file in Google Apps Script and rename it to "SheetsJS." Insert the data found in the SheetsJS.js file from this folder.

## Author

These scripts were developed by [Rodrigo Frutuoso](https://github.com/rodfrutuoso) to facilitate automation tasks in Google Sheets spreadsheets.

---



