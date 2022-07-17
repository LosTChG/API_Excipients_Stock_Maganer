function createSpreadsheetOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('alAbrir')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
}
