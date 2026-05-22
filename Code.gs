function doGet(){
  return HtmlService.createHtmlOutput('<script>location.href="https://formulario-arquitetos.vercel.app"</script>');
}

function doPost(e){
  try {
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
  } catch(err) {}

  return HtmlService.createHtmlOutput('<script>location.href="https://formulario-arquitetos.vercel.app?ok=1"</script>');
}
