// Webhook
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.token) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: "erro", mensagem: "token não informado" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

/*
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Token");
    sheet.getRange("A1:B1").setValues([[data.token, new Date()]]);
*/

    // Salva o token e a data no Script Properties
    const props = PropertiesService.getScriptProperties();
    props.setProperty("TOKEN", data.token);
    props.setProperty("ATUALIZADO_EM", new Date().toISOString());

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "ok",
        tokenSalvo: data.token,
        atualizadoEm: props.getProperty("ATUALIZADO_EM")
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "erro", mensagem: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Report
function report(){
  
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  //const token = ss.getSheetByName('Token').getRange(1,1).getValue()
  const token = PropertiesService.getScriptProperties().getProperty("TOKEN");


  const finalDate = new Date();
  const startDate = new Date();
  startDate.setDate(finalDate.getDate() - 30);

  const base64 = requestNewReport(token, startDate.toISOString(), finalDate.toISOString())
  const data = convertBase64toSheets(base64)
  const values = parseReport(data)
  
  const sheet = ss.getSheetByName("Movimentações")
  const last = sheet.getLastRow()
  if (last > 3) sheet.deleteRows(4, last - 1)
  
  sheet.getRange(3, 1, values.length, values[0].length).setValues(values); 
}

function requestNewReport(token, startDate, finalDate){
  const url = 'https://novaapirelatorios.verapp.com.br/v3/ReportMovimentacoesEstoque/excel';

  const payload = {
    ascending: false,
    orderBy: "Data",
    genero: "04 – Produto Acabado",
    tipo: "",
    periodoData: 11,
    dataInicial: startDate,
    dataFinal: finalDate,
    deposito: "",
    depositoID: "",
    responsavelMovimentacao: "",
    codigoEAN: ""
  };

  const options = {
    method: 'post',
    headers: {
      'authorization': `Bearer ${token}`,
      'Accept': "application/octet-stream",
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);

  if (res.getResponseCode() !== 200) {
    Logger.log("Erro na requisição: " + res.getContentText());
    return;
  }

  const json = JSON.parse(res.getContentText())
  return json.Data.split(",")[1];
}

function convertBase64toSheets(base64){
  
  const content = Utilities.base64Decode(base64);

  const blob = Utilities.newBlob(content, MimeType.MICROSOFT_EXCEL, "tmp.xlsx");
  const schema = { title: "TEMP_IMPORT", mimeType: MimeType.GOOGLE_SHEETS }

  const tempFile = Drive.Files.create(schema, blob, { convert: true } );
  
  const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
  const tempSheet = tempSpreadsheet.getSheets()[0];

  return tempSheet.getDataRange().getValues();
}

function parseReport(data){
  for(let i = 3; i < data.length; i++){
    if(data[i][0] == ''){
      data[i][0] = data[i-1][0]
      data[i][1] = data[i-1][1]
    }
  }
  return data.filter(v => v[2] == 'Entrada')
}
