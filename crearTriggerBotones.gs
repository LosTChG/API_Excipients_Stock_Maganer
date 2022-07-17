function createSpreadsheetEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('alPulsarBoton')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
}
