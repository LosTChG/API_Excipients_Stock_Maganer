function createSpreadsheetEditTrigger() {
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  const SS = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('alPulsarBoton')
      .forSpreadsheet(SS)
      .onEdit()
      .create();

  var datos_sheet = SS.getSheets()[1];
  datos_sheet.getRange("J19").setValue("Ficha habilitada");
}
