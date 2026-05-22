function doGet(){
  return HtmlService.createHtmlOutput('<script>location.href="https://formulario-arquitetos.vercel.app"</script>');
}

function doPost(e){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Respostas");
  if (!sheet) {
    sheet = ss.insertSheet("Respostas");
  }

  const headers = ["Data", "Nome", "Email", "WhatsApp", "Área de Atuação", "Interesse"];

  if (!sheet.getRange("A1").getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nome,
    e.parameter.email,
    e.parameter.whatsapp,
    e.parameter.area,
    e.parameter.interesse
  ]);

  return HtmlService.createHtmlOutput('<script>location.href="https://formulario-arquitetos.vercel.app?ok=1"</script>');
}
