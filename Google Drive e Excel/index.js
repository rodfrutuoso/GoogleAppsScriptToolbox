/**
 * Cria um menu na planilha para acionar os scripts
 * 
 * Create a menu in the spreadsheet to trigger the scripts
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Scripts')
        .addItem('Atualizar Arquivos', 'getFolderTree')
        .addSeparator()
        .addItem('Extrair Dados', 'PegarAbas')
        .addToUi();
}

/**
 * Através da Api do Google Drive e da Api do Google Sheets, lista todos os arquivos e pastas da 
 * pasta do id da variáveç folderId e coloca essa relação da aba RELAÇÃO ATUAL
 * 
 * Using the Google Drive Api and the Google Sheets Api, list all the files and folders
 * in the folder ID of the variable folderId and place this relationship in the RELAÇÃO ATUAL tab
 */
function getFolderTree() {
    var folderId = "######################";    // insira o link da pasta do Google Drive aqui // insert Google Drive folder link here
    var abaListarArquivos = "RELAÇÃO ATUAL!A"; // aba que será listado os arquivos da pasta // sheet that will get the files list in the folder
    var listAll = true;
    var todosdados = [];

    let options = {
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        },
        supportsAllDrives: true,
    };

    var folder = Drive.Files.get(folderId, options);
    Logger.log(folder.title)
    getChildFolders(folder.title, folder, listAll, todosdados, 1);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getId()
    //sheet.clear();
    Logger.log(todosdados);

    var vio = { valueInputOption: 'USER_ENTERED' }
    var resource = { values: todosdados }

    Sheets.Spreadsheets.Values.clear({}, sheet, abaListarArquivos+"1:E")
    let titulos = [["Full Path", "Level", "Type", "Date", "URL"]]
    var resource2 = { values: titulos }
    Sheets.Spreadsheets.Values.update(resource2, sheet, abaListarArquivos + "1", vio)
    Sheets.Spreadsheets.Values.update(resource, sheet, abaListarArquivos+"2", vio)


}

/**
 * Função recursiva da getFolderTree() para listar pastas
 */
function getChildFolders(parentName, parent, listAll, todosdados, level) {
    var pageToken = '';
    var folderss = [];
    do {
        var folders = Drive.Files.list({
            q: "mimeType='application/vnd.google-apps.folder' and trashed=false and '" + parent.id + "' in parents",
            fields: "items(createdDate,id,modifiedDate,title,alternateLink,mimeType),nextPageToken",
            supportsAllDrives: true,
            headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
            includeItemsFromAllDrives: true,
            maxResults: 1000,
            pageToken: pageToken
        });

        folderss = folderss.concat(folders.items);
        pageToken = folders.nextPageToken;
    } while (pageToken);

    if (level === 1) {
        getChildFiles(parentName, parent, todosdados, 1);
    }

    for (var i = 0; i < folderss.length; i++) { //Utilities.formatDate(file.createdDate,"GMT-3","dd/MM/yyyy")
        var folder = folderss[i];
        var data = [parentName + "/" + folder.title, level, "PASTA", Utilities.formatDate(new Date(folder.createdDate), "GMT-3", "dd/MM/yyyy"), folder.alternateLink];
        todosdados.push(data);

        getChildFiles(parentName + "/" + folder.title, folder, todosdados, level + 1);
        getChildFolders(parentName + "/" + folder.title, folder, listAll, todosdados, level + 1)
    }
}

/**
 * Função recursiva da getFolderTree() para listar arquivos
 */
function getChildFiles(parentName, parent, todosdados, level) {
    var pageToken = '';
    var files = [];

    do {
        var result = Drive.Files.list({
            q: "mimeType!='application/vnd.google-apps.folder' and trashed=false and '" + parent.id + "' in parents",
            fields: "items(createdDate,id,modifiedDate,title,alternateLink,mimeType),nextPageToken",
            supportsAllDrives: true,
            headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
            includeItemsFromAllDrives: true,
            maxResults: 1000,
            pageToken: pageToken
        });
        files = files.concat(result.items);
        pageToken = result.nextPageToken;
    } while (pageToken);

    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var data = [parentName + "/" + file.title, level, "ARQUIVO", Utilities.formatDate(new Date(file.createdDate), "GMT-3", "dd/MM/yyyy"), file.alternateLink];
        todosdados.push(data);
    }
}

/**
 * Pega os ID's dos links dos arquivos em excel 
 */
function PegarAbas() {

    let id_sheets_copia = SpreadsheetApp.getActiveSpreadsheet().getId()
    var vio = { valueInputOption: 'USER_ENTERED' }

    // Pegas os Ids das planilhas em exel e armazena em "ids" // Take the Ids of the spreadsheets in exel and store them in "ids"
    let range_sheets_copia = "RELAÇÃO ATUAL!L2:L" // range que possui o REGEXEXTRACT com os IDs das planilhas // range that has the REGEXEXTRACT with the worksheet IDs
    let range_sheets_cola = "CONTEÚDOS!A1:J" // rangue que será inserido o conteúdo dos arquivos em excel // range where the contents of the excel files will be inserted
    let range_sheets_cola_AUX = "CONTEÚDOS!A" // rangue auxiliar para inserir os dados // auxiliary slot for entering data
    let abaExcel = "DESPESAS DE VIAGEM" // aba que deseja extrair as informações do arquivo em excel // sheet that you want to extract information from the excel file
    var values_cola = Sheets.Spreadsheets.Values.get(id_sheets_copia, range_sheets_copia).values.filter(String);
    var expReg = /\/d\/(.*)\//
    var ids = []
    var aux
    values_cola.map(element => {
        aux = element[0].match(expReg)
        ids.push(aux[1])
    });

    Logger.log(ids)
    Sheets.Spreadsheets.Values.clear({}, id_sheets_copia, range_sheets_cola)

    //Abre o excel através da bibliotedca SheetJS e insere os dados da aba CONTEÚDOS na aba DESPESAS DE VIAGEM
    ids.map(element => {
        try {
            var ultimaLinha = Sheets.Spreadsheets.Values.get(id_sheets_copia, range_sheets_cola).values.length
        } catch (error) {
            var ultimaLinha = 1
        }

        //Obtém o arquivo através do ID // Gets the file using ID
        const data = new Uint8Array(DriveApp.getFileById(element).getBlob().getBytes());
        const book = XLSX.read(data, { type: "array" });
        //Converte o excel em CSV e pega os dados //Convert excel to CSV and get the data
        const csv = XLSX.utils.sheet_to_csv(book.Sheets[abaExcel]);
        const values = Utilities.parseCsv(csv);

        // Insre os dados na planilha // insert the data into the sheet
        var resource = { values: values }
        Sheets.Spreadsheets.Values.update(resource, id_sheets_copia, range_sheets_cola_AUX + ultimaLinha, vio)

    });

}

