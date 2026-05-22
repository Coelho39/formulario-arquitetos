function doGet(){
  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas");

  if(!sheet.getRange("A1").getValue()){
    sheet.getRange("1:1").setValues([[
      "Data", "Nome", "Email", "WhatsApp", "Área de Atuação", "Interesse"
    ]]);
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nome,
    e.parameter.email,
    e.parameter.whatsapp,
    e.parameter.area,
    e.parameter.interesse
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({status: "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}
