/**
 * Importa o Relatório de Chamado do desk através do link do URL Power BI. 
 * Para isso, é utilizada classe UrlFetchApp que consulta o HTML da página.
 * 
 * Imports the Ticket Report from the desk via the Power BI URL link.
 * To do this, use the UrlFetchApp class, which queries the HTML of the page.
 */
function importDeskHTML() {

    var url = 'https://***.desk.ms/Relatorios/powerbi?token=******' //url do relatório // Report url 
    var html = '<table' + UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText().replace(/(\r\n|\n|\r|\t|  )/gm, "").match(/(?<=\<table).*(?=\<\/table)/g) + '</table>';
    var trs = [...html.matchAll(/<tr[\s\S\w]+?<\/tr>/g)];
    var data = [];
    for (var i = 0; i < trs.length; i++) {
        var tds = [...trs[i][0].matchAll(/<(td|th)[\s\S\w]+?<\/(td|th)>/g)];
        var prov = [];
        for (var j = 0; j < tds.length; j++) {
            donnee = tds[j][0].match(/(?<=\>).*(?=\<\/)/g)[0];
            prov.push(stripTags(donnee));
        }
        data.push(prov);
    }

    return data
}

/**
 * Verifica quais sãos os novos chamados que ainda não estão na planilha e os insere utilizando a API do Google Sheets
 * 
 * Checks which new tickets are not yet in the spreadsheet and inserts them using the Google Sheets API
 */
function filtroChamdosNovos() {
    let abaRange = "SHEET!A2:A" // Range da coluna que tem os números de chamado // Range of the column with ticket numbers
    let abaRangeAux = "SHEET!A"// Mesmo range porém sem o número da linha para inserir os dados posteriormente // Same range but without the line number to insert the data later
    
    //pega os chamados que já existem na planilha // get the tickes that already exist in the spreadsheet
    var id = SpreadsheetApp.getActiveSpreadsheet().getId()
    var dadosAba = Sheets.Spreadsheets.Values.get(id, abaRange).values
    dadosAba = dadosAba.map(function (row) { return row[0] })
    var ultimalinha = dadosAba.length + 7
    Logger.log(dadosAba)
    var dados = importDeskHTML()

    //ajusta o vetor dos dados // ajust data array
    var colVerificacao = dados.map(function (row) { return row[0] })
    Logger.log(colVerificacao)

    //verifica quais chamados novos ainda não foram inseridos na planilha e os insere // filter new tickets
    var dadosFinais = dados.filter(function (row) {
        return !dadosAba.includes(row[0])
    })

    Logger.log(dadosFinais)
    var vio = { valueInputOption: 'USER_ENTERED' }
    var id = SpreadsheetApp.getActiveSpreadsheet().getId()
    var resource1 = { values: dadosFinais.slice(1) }
    var range_sheets_cola = abaRangeAux + ultimalinha


    //insere os dados na planilha //insert new tickets in the spreadsheet
    Sheets.Spreadsheets.Values.update(resource1, id, range_sheets_cola, vio)
}

/**
 * No google sheets há um limite de 5000 caracteres por célula.
 * Essa função impede que tamanho da string do vetor de 'dados' ultrapasse esse valor.
 * 
 * In google sheets there is a limit of 5000 characters per cell.
 * This function prevents the string size of the 'dados' from exceeding this value.
 */
function stripTags(body) {
    var regex = /(<([^>]+)>)/ig;
    return body.replace(regex, "").slice(0, 4998);
}