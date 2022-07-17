function alPulsarBoton(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();

  if (sheetName == "MP ya existente") {
    if (e.range.getA1Notation() == "C17") {
      if (e.value == "Crear Lote MP") {
        botonCrearLoteMP();
      } else if (e.value == "Reset") {
        botonResetMP();
      }
      
      e.range.setValue("");
    }
  } else if (sheetName == "Nueva MP") {
    if (e.range.getA1Notation() == "C21") {
      if (e.value == "Crear Nueva MP") {
        botonCrearNuevaMP()
      } else if (e.value == "Reset") {
        botonResetNuevaMP();
      }
      
      e.range.setValue("");
    }
  } else if (sheetName == "Resumen MPs") {
    if (e.range.getA1Notation() == "I2") {
      if (e.value == "Refrescar Datos") {
        botonActualizarResumen();
      } else if (e.value == "Refrescar Etiquetas (simple)") {
        botonActualizarEtiquetasMPs(false,false);
      } else if (e.value == "Refrescar Etiquetas (con código de envase)") {
        botonActualizarEtiquetasMPs(false,true);
      } else if (e.value == "Refrescar Etiquetas (con código de envase y estado)") {
        botonActualizarEtiquetasMPs(true,true);
      } else if (e.value == "Refrescar Datos de Todos") {
        botonActualizarResumenTODOS();
      }
      
      e.range.setValue("");
    } else if (e.range.getA1Notation() == "H2") {
      sheet.getRange("J2").setValue("");
    }
  }
}

function alAbrir() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[1];
  var sheetResumen = SS.getSheets()[2];

  var numeroMPs = parseInt(sheetResumen.getRange("D2").getValue());

  var cell = sheet.getRange("C2");
  cell.setValue(numeroMPs+1);

  resetMPYaExistente();
};

function resetMPYaExistente() {
  // Desplegable MP ya existente
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[0];
  var sheetResumen = SS.getSheets()[2];

  cell = sheet.getRange("C2");
  cell.setValue(sheetResumen.getRange("A5").getValue());

  var numero_filas = 4+parseInt(sheetResumen.getRange("D2").getValue());
  var rango = sheetResumen.getRange("A5:A"+numero_filas);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rango, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
}

function existeCarpeta(nombre_carpeta) {
  var id_directorio = "1XpHhcGvTevBEYWlKFnpxvLVTtd6lWK8c";
  
  try {
    // Get folder by id
    var parentFolder = DriveApp.getFolderById(id_directorio);
       
    // Get folders en ese folder
    var childFolders = parentFolder.getFolders();

    var carpeta = null;
    while (childFolders.hasNext()) {
      carpeta = childFolders.next();
      if (carpeta.getName() == nombre_carpeta) {
        break;
      } else {
        carpeta = null;
      }
    } 

    return carpeta;
  } catch (e) {
    Logger.log(e.toString());
  }
};

function eliminarSubCarpetas(carpeta_padre) {  
  // Get folders en ese folder
  var childFolders = carpeta_padre.getFolders();

  while (childFolders.hasNext()) {
    var carpeta = childFolders.next();
    carpeta_padre.removeFolder(carpeta);
  } 
};

function copiarContenidoCarpeta(carpeta_contenedora, carpeta_destino) {
  var filesIterator = carpeta_contenedora.getFiles();

  while (filesIterator.hasNext()) {
    var file = filesIterator.next();
    file.makeCopy(file.getName(), carpeta_destino);
  }
}

function buscarFilaEnResumen(containingValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets();
  var dataRange = sheet[2].getRange(5,2,1002);
  var values = dataRange.getValues();

  var outRow;

  for (var i = 0; i < values.length; i++)
  {
    if (values[i] == containingValue)
    {
      outRow = 5+i;
      break;
    }
  }

  return outRow;
}

function botonActualizarResumen() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResumen = SS.getSheets()[2];

  sheetResumen.getRange("J2").setValue("REFRESCANDO...");

  var codigo = parseInt(sheetResumen.getRange("H2").getValue());
  if(actualizarResumenMPs(codigo)) {
    sheetResumen.getRange("J2").setValue("");
  } else {
    sheetResumen.getRange("J2").setValue("NO EXISTE ESE CÓDIGO");

    // Considerar como eliminado (si existe)
    var fila = buscarFilaEnResumen(codigo);
    if (fila != undefined) {
      sheetResumen.getRange(fila,2).setValue("");
      sheetResumen.getRange(fila,3).setValue("");
      sheetResumen.getRange(fila,4).setValue("");
      sheetResumen.getRange(fila,5).setValue("");
      sheetResumen.getRange(fila,6).setValue("");
      sheetResumen.getRange(fila,7).setValue("");
      sheetResumen.getRange(fila,8).setValue("");
      sheetResumen.getRange(fila,9).setValue("");
      sheetResumen.getRange(fila,10).setValue("");
      sheetResumen.getRange(fila,11).setValue("");
      sheetResumen.getRange(fila,12).setValue("");

      alAbrir();
    }
  }
}

function botonActualizarResumenTODOS() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResumen = SS.getSheets()[2];

  sheetResumen.getRange("J2").setValue("REFRESCANDO... LLEVARÁ BASTANTE RATO");

  var codigo = parseInt(sheetResumen.getRange("H2").getValue());
  var max_codigo = sheetResumen.getRange("C2").getValue();

  for (var i = 1;i<=max_codigo;i++) {
    actualizarResumenMPs(1);
  }
  
  sheetResumen.getRange("J2").setValue("");
  alAbrir();
}

// Refresca Etiquetas MP
function botonActualizarEtiquetasMPs(poner_estado,poner_envase) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResumen = SS.getSheets()[2];

  sheetResumen.getRange("J2").setValue("REFRESCANDO...");

  var codigo = parseInt(sheetResumen.getRange("H2").getValue());

  var carpeta = existeCarpeta(codigo);
  if(carpeta != null) {
    actualizarEtiquetasMPs(carpeta,poner_estado,poner_envase);
    sheetResumen.getRange("J2").setValue("");
  } else {
    sheetResumen.getRange("J2").setValue("NO EXISTE ESE CÓDIGO");

    // Considerar como eliminado (si existe)
    var fila = buscarFilaEnResumen(codigo);
    if (fila != undefined) {
      sheetResumen.getRange(fila,10).setValue("ELIMINADO");
    }
  }
}

function actualizarEtiquetasMPs(carpeta_padre, poner_estado, poner_envase) {
  var documento = carpeta_padre.getFilesByName("Etiquetas 21").next();
  var body = DocumentApp.openById(documento.getId()).getBody();
  var tabla = body.findElement(DocumentApp.ElementType.TABLE).getElement().asTable();

  // Get folders en ese folder
  var childFolders = carpeta_padre.getFolders();

  var row = tabla.getRow(0);
  var numero_row = 0;
  var numero_celda_row = 0;
  while(childFolders.hasNext()) {
    if (numero_celda_row > 8) {
      numero_row += 1;
      row = tabla.getRow(numero_row);
      numero_celda_row = 0;
    }
    var celda = row.getCell(numero_celda_row);
    
    // QR
    celda.clear();
    var carpeta = childFolders.next();
    var archivos = carpeta.getFiles();
    var qr_encontrado = false;
    while(archivos.hasNext()) {
      var archivo_QR = archivos.next();
      if(archivo_QR.getName() == "QR envase.png") {
        celda.insertImage(0,archivo_QR.getBlob()).setWidth(120).setHeight(120);
        celda.removeChild(celda.getChild(1));
        qr_encontrado = true;
        break;
      }
    }
    if(!qr_encontrado) {
      celda.setText("¡Falta generar QR!");
    }
    
    numero_celda_row += 1;
    celda = celda.getNextSibling().asTableCell();
    
    // Info
    celda.clear();
    var archivo_ficha = carpeta.getFilesByName("Ficha MP").next();
    var excel_ficha = SpreadsheetApp.open(archivo_ficha);
    var datos_sheet = excel_ficha.getSheets()[1];
    var codigo = datos_sheet.getRange("C2").getValue();
    if(poner_estado) {
      var estado = datos_sheet.getRange("C19").getValue();
      var caducado = datos_sheet.getRange("D19").getValue();
      if (caducado == "CADUCADO") {
        estado = caducado;
      }
      var estado_paragrafo = celda.appendParagraph(estado);
      estado_paragrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      estado_paragrafo.editAsText().setBold(true);
      if (estado == "APROBADO") {
        estado_paragrafo.editAsText().setBackgroundColor("#b7e1cd");
      } else if (estado == "CUARENTENA") {
        estado_paragrafo.editAsText().setBackgroundColor("#fce8b2");
      } else if (estado == "RETIRADO" || estado == "CADUCADO") {
        estado_paragrafo.editAsText().setBackgroundColor("#f4c7c3");
      } else if (estado == "ELIMINADO") {
        estado_paragrafo.editAsText().setStrikethrough(true);
      }
      
      celda.appendParagraph("").editAsText().setBold(false).setBackgroundColor("#ffffff");
    }
    celda.appendParagraph("MP-"+codigo).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    if (poner_envase) {
      celda.appendParagraph("");

      var codigo_envase = datos_sheet.getRange("C3").getValue();
      celda.appendParagraph(codigo_envase).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    celda.removeChild(celda.getChild(0));

    numero_celda_row += 2;
  }
}

function actualizarResumenMPs(codigo) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResumen = SS.getSheets()[2];

  // Coger datos que nos interesan
  var carpeta_padre = existeCarpeta(codigo);
  if (carpeta_padre == null) {
    return false;
  }
  var childFolders = carpeta_padre.getFolders();

  // Coger fila del Resumen
  var fila = buscarFilaEnResumen(codigo);
  if (fila == undefined) {
    var numero_filas = 4+parseInt(sheetResumen.getRange("D2").getValue());
    fila = numero_filas+1;
  }

  var nombre;
  var cas;
  var proveedor;
  var referencia;

  var ultima_caducidad = transformarFecha(sheetResumen.getRange(fila,7).getValue());
  var stock_disponible = 0.0;
  var estado;

  var primera_vez = true;
  while (childFolders.hasNext()) {
    var carpeta = childFolders.next();
    var archivo_ficha = carpeta.getFilesByName("Ficha MP").next();
    var excel_ficha = SpreadsheetApp.open(archivo_ficha);
    var datos_sheet = excel_ficha.getSheets()[1];
    var consumos_sheet = excel_ficha.getSheets()[2];

    if (primera_vez) {
      nombre = datos_sheet.getRange("C5").getValue();
      cas = datos_sheet.getRange("C6").getValue();
      proveedor = datos_sheet.getRange("C8").getValue();
      referencia = datos_sheet.getRange("C9").getValue();

      // Colocar los datos en la sheetResumen
      sheetResumen.getRange(fila,2).setValue(codigo);
      sheetResumen.getRange(fila,3).setValue(nombre);
      sheetResumen.getRange(fila,4).setValue(cas);
      sheetResumen.getRange(fila,5).setValue(proveedor);
      sheetResumen.getRange(fila,6).setValue(referencia);

      primera_vez = false;
    }

    var fecha = transformarFecha(datos_sheet.getRange("C17").getValue());
    if (ultima_caducidad < fecha) {
      ultima_caducidad = fecha;
    }

    var estado_pre = datos_sheet.getRange("C19").getValue();
    var caducado = datos_sheet.getRange("D19").getValue();
    if (caducado == "" && (estado_pre == "APROBADO" || estado_pre == "CUARENTENA")) {
      stock_disponible += parseFloat(consumos_sheet.getRange("F3").getValue());
    }
    
    if (estado_pre == "APROBADO") {
      estado = estado_pre;
    } else if (estado_pre == "CUARENTENA" && estado != "APROBADO") {
      estado = estado_pre;
    } else if (estado_pre == "RETIRADO" && estado != "APROBADO" && estado != "CUARENTENA") {
      estado = estado_pre;
    } else if (estado_pre == "ELIMINADO" && estado != "APROBADO" && estado != "CUARENTENA" && estado != "RETIRADO") {
      estado = estado_pre;
    }
  }

  sheetResumen.getRange(fila,7).setValue(ultima_caducidad);
  sheetResumen.getRange(fila,8).setValue(stock_disponible);
  sheetResumen.getRange(fila,10).setValue(estado);
    // Link carpeta
  sheetResumen.getRange(fila,12).setValue(carpeta_padre.getUrl());

  return true;
}



function transformarFecha(fecha_transformar) {
  if (typeof fecha_transformar == "string") {
    var day = +fecha_transformar.substring(0, 2);
    var month = +fecha_transformar.substring(3, 5);
    var year = +fecha_transformar.substring(6, 10);

    return new Date(year, month - 1, day);
  } else {
    return fecha_transformar;
  }

}

function botonCrearLoteMP() {
  // Datos origen
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[0];
  var sheetResumen = SS.getSheets()[2];

  sheet.getRange("D17").setValue("CREANDO NUEVO LOTE...");

  var codigo = sheet.getRange("C2").getValue().substring(0,1);

  var fila = buscarFilaEnResumen(codigo);

  var nombre = sheetResumen.getRange(fila,3).getValue();
  var cas = sheetResumen.getRange(fila,4).getValue();
  var proveedor = sheetResumen.getRange(fila,5).getValue();
  var referencia = sheetResumen.getRange(fila,6).getValue();

  var lote_proveedor = sheet.getRange("C4").getValue();
  var riqueza = sheet.getRange("C5").getValue();

  var peso_por_envase = sheet.getRange("C7").getValue();
  var cantidad_envases = parseInt(sheet.getRange("C8").getValue());
  var modo_inetiquetable = sheet.getRange("C9").getValue();

  var fecha_recepcion = sheet.getRange("C11").getValue();
  var fecha_reanalisis = sheet.getRange("C12").getValue();

  var observaciones = sheet.getRange("C14").getValue();


  crearArchivos(codigo, nombre, cas, proveedor, referencia, lote_proveedor, riqueza, peso_por_envase, cantidad_envases, modo_inetiquetable, fecha_recepcion, fecha_reanalisis, observaciones, false);
    
  actualizarResumenMPs(codigo);

  sheet.getRange("D17").setValue("");
}

function botonCrearNuevaMP() {
  // Datos origen
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[1];

  sheet.getRange("D21").setValue("CREANDO NUEVA MP...");

  var codigo = sheet.getRange("C2").getValue();

  var nombre = sheet.getRange("C4").getValue();
  var cas = sheet.getRange("C5").getValue();

  var proveedor = sheet.getRange("C7").getValue();
  var referencia = sheet.getRange("C8").getValue();
  var lote_proveedor = sheet.getRange("C9").getValue();
  var riqueza = sheet.getRange("C10").getValue();

  var peso_por_envase = sheet.getRange("C12").getValue();
  var cantidad_envases = parseInt(sheet.getRange("C13").getValue());
  var modo_inetiquetable = sheet.getRange("C14").getValue();

  var fecha_recepcion = sheet.getRange("C16").getValue();
  var fecha_reanalisis = sheet.getRange("C17").getValue();

  var observaciones = sheet.getRange("C19").getValue();


  crearArchivos(codigo, nombre, cas, proveedor, referencia, lote_proveedor, riqueza, peso_por_envase, cantidad_envases, modo_inetiquetable, fecha_recepcion, fecha_reanalisis, observaciones, true);

  actualizarResumenMPs(codigo);

  sheet.getRange("D21").setValue("");
};

function crearArchivos(codigo, nombre, cas, proveedor, referencia, lote_proveedor, riqueza, peso_por_envase, cantidad_envases, modo_inetiquetable, fecha_recepcion, fecha_reanalisis, observaciones, nuevaMP) {
  // Crear nuevos archivos   
  var id_directorio = "1XpHhcGvTevBEYWlKFnpxvLVTtd6lWK8c";
  var id_dir_etiquetas = "1hE-Dq_lQiUpoMJURsn8Hu1HGAlr2iUq1";
  var parentFolder = DriveApp.getFolderById(id_directorio);

  var carpeta_objetivo = existeCarpeta(codigo);

  var envases_existentes = 0;
  if (nuevaMP) {
    if (carpeta_objetivo != null) {
      eliminarSubCarpetas(carpeta_objetivo);
    } else {
      carpeta_objetivo = parentFolder.createFolder(codigo);

      // Etiquetas
      var carpeta_plantilla_etiq = DriveApp.getFolderById(id_dir_etiquetas);
      copiarContenidoCarpeta(carpeta_plantilla_etiq, carpeta_objetivo);
    }
  } else {
    var hoy = new Date();
    var carpetas_existentes = carpeta_objetivo.getFolders();
    while(carpetas_existentes.hasNext()) {
      var carp = carpetas_existentes.next();
      if (carp.getName().startsWith(codigo+"-"+hoy.getFullYear())) {      
        envases_existentes += 1;
      }
    }
  }


  for (let i = 1; i <= cantidad_envases; i++) {
    var carpeta_plantilla = DriveApp.getFolderById("1xye0F8tNd3p7eGJvlxOh7SKGJGejPD19");
    
    var hoy = new Date();
    var canti_enva = envases_existentes+i;
    var nombre_envase = hoy.getFullYear()+"_"+canti_enva;
    var nombre_carpeta_envase = carpeta_objetivo.getName()+"-"+nombre_envase;

    var carpeta_envase = carpeta_objetivo.createFolder(nombre_carpeta_envase);
    
    copiarContenidoCarpeta(carpeta_plantilla, carpeta_envase);
    var archivo_ficha = carpeta_envase.getFilesByName("Ficha MP").next();
    var excel_ficha = SpreadsheetApp.open(archivo_ficha);
  

    // Copiar datos a la ficha

    var datos_sheet = excel_ficha.getSheets()[1];
    datos_sheet.getRange("C2").setValue(codigo);
    datos_sheet.getRange("C3").setValue(nombre_envase);
    datos_sheet.getRange("C5").setValue(nombre);
    datos_sheet.getRange("C6").setValue(cas);
    datos_sheet.getRange("C8").setValue(proveedor);
    datos_sheet.getRange("C9").setValue(referencia);
    datos_sheet.getRange("C10").setValue(lote_proveedor);
    datos_sheet.getRange("C11").setValue(riqueza);
    datos_sheet.getRange("C13").setValue(peso_por_envase);
    if (modo_inetiquetable == "SÍ") {
      datos_sheet.getRange("C14").setValue(cantidad_envases);
    }
    datos_sheet.getRange("C16").setValue(fecha_recepcion);
    datos_sheet.getRange("C17").setValue(fecha_reanalisis);
    datos_sheet.getRange("C21").setValue(observaciones);

    // Modificar .bat
    var archivo_bat = carpeta_envase.getFilesByName("generarQR.bat").next();
    var contenido_bat = archivo_bat.getBlob().getDataAsString();
    var contenido_bat_anadir = " \""+archivo_ficha.getUrl()+"\" \"QR envase.png\"";
    archivo_bat.setContent(contenido_bat+contenido_bat_anadir);
  }
  
  // Actualizar número
  // alAbrir();
}

function botonResetNuevaMP() {
  alAbrir();

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[1];

  sheet.getRange("C4").setValue("");
  sheet.getRange("C5").setValue("");

  sheet.getRange("C7").setValue("");
  sheet.getRange("C8").setValue("");
  sheet.getRange("C9").setValue("");
  sheet.getRange("C10").setValue("");

  sheet.getRange("C12").setValue("");
  sheet.getRange("C13").setValue("");
  sheet.getRange("C14").setValue("NO");

  sheet.getRange("C16").setValue("");
  sheet.getRange("C17").setValue("");

  sheet.getRange("C19").setValue("");
};

function botonResetMP() {
  resetMPYaExistente();

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheets()[0];

  sheet.getRange("C4").setValue("");
  sheet.getRange("C5").setValue("");

  sheet.getRange("C7").setValue("");
  sheet.getRange("C8").setValue("");
  sheet.getRange("C9").setValue("");

  sheet.getRange("C11").setValue("");
  sheet.getRange("C13").setValue("");

  sheet.getRange("C14").setValue("");
};
