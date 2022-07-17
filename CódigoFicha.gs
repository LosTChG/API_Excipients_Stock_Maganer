function alPulsarBoton(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();

  if (sheetName == "Nuevo consumo") {
    if (e.range.getA1Notation() == "C14") {
      if (e.value == "Registrar consumo") {
        botonRegistrarConsumo();
      } else if (e.value == "Reset") {
        botonResetearConsumo();
      }
      
      e.range.setValue("");
    }
  }
}

function botonRegistrarConsumo() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var nuevo_cons_sheet = SS.getSheets()[0];
  var datos_sheet = SS.getSheets()[1];
  var consumos_sheet = SS.getSheets()[2];

  nuevo_cons_sheet.getRange("D14").setValue("REGISTRANDO NUEVO CONSUMO...");

  var destino = nuevo_cons_sheet.getRange("C2").getValue();
  var lote_destino = nuevo_cons_sheet.getRange("C3").getValue();
  var peso_antes = nuevo_cons_sheet.getRange("C5").getValue();
  var peso_despues = nuevo_cons_sheet.getRange("C6").getValue();
  var cantidad_gastada = nuevo_cons_sheet.getRange("C7").getValue();
  var observaciones = nuevo_cons_sheet.getRange("C9").getValue();
  var fecha = nuevo_cons_sheet.getRange("C11").getValue();
  var firma = nuevo_cons_sheet.getRange("C12").getValue();

  if (datos_sheet.getRange("D19").getValue() == "CADUCADO") {
    nuevo_cons_sheet.getRange("D14").setValue("ATENCIÓN: El bote está CADUCADO");
    //var htmlOutput = HtmlService.createHtmlOutput('ATENCIÓN: El bote está CADUCADO');
    //SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bote CADUCADO');
  } else if (datos_sheet.getRange("C19").getValue() != "APROBADO") {
    nuevo_cons_sheet.getRange("D14").setValue("ATENCIÓN: El bote está "+datos_sheet.getRange("C19").getValue());
    //var htmlOutput = HtmlService.createHtmlOutput('ATENCIÓN: El bote está '+datos_sheet.getRange("C19").getValue());
    //SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bote '+datos_sheet.getRange("C19").getValue());
  } else {
    if (firma == "" || fecha == "" || destino == "") {
      nuevo_cons_sheet.getRange("D14").setValue("Faltan datos obligatorios por introducir!");
      //var htmlOutput = HtmlService.createHtmlOutput('Faltan datos obligatorios por introducir!');
      //SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Faltan datos');
    } else {
      var n_consumos = parseInt(consumos_sheet.getRange("B3").getValue());
      var fila = n_consumos+5;

      consumos_sheet.getRange(fila,2).setValue(destino);
      consumos_sheet.getRange(fila,3).setValue(lote_destino);
      consumos_sheet.getRange(fila,4).setValue(peso_antes);
      consumos_sheet.getRange(fila,5).setValue(peso_despues);
      consumos_sheet.getRange(fila,6).setValue(cantidad_gastada);
      if (parseFloat(consumos_sheet.getRange("F3").getValue()) <= 0.0) {
        var htmlOutput = HtmlService
          .createHtmlOutput('ATENCIÓN: En teoría se ha acabado el contenido del bote');
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Se ha acabado el bote');
        datos_sheet.getRange("C19").setValue("RETIRADO");
      }
      consumos_sheet.getRange(fila,7).setValue(observaciones);
      consumos_sheet.getRange(fila,8).setValue(fecha);
      consumos_sheet.getRange(fila,9).setValue(firma);

      var codigo = datos_sheet.getRange("C2").getValue();
      var url = "https://script.google.com/a/macros/unav.es/s/AKfycbyAx0e4uafGfhP7yhKwyfSygjPk-U1Rw-q-2LEpor6vOX6hGc1Ewk8iDHFZaW2bEhyg/exec?refrescar="+codigo;

      nuevo_cons_sheet.getRange("D14").setValue("Actualizar Resumen MP:");
      nuevo_cons_sheet.getRange("E14").setValue(url);
      //var htmlOutput = HtmlService.createHtmlOutput('Hacer click <a href="'+url+'" target="_blank">aquí</a> para actualizar el Resumen de MP');
      //SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Actualizar Resumen MP');
    }
  }
}

function botonResetearConsumo() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var nuevo_cons_sheet = SS.getSheets()[0];

  nuevo_cons_sheet.getRange("C2").setValue("");
  nuevo_cons_sheet.getRange("C3").setValue("");
  nuevo_cons_sheet.getRange("C5").setValue("");
  nuevo_cons_sheet.getRange("C6").setValue("");
  nuevo_cons_sheet.getRange("C7").setValue("=C5-C6");
  nuevo_cons_sheet.getRange("C9").setValue("");
  nuevo_cons_sheet.getRange("C11").setValue("=TODAY()");
  nuevo_cons_sheet.getRange("C12").setValue("");
  nuevo_cons_sheet.getRange("D14").setValue("");
  nuevo_cons_sheet.getRange("E14").setValue("");
}
