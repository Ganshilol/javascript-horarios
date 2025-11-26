function consultarFechaEnRango() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hojaConsulta = planilla.getSheetByName("consulta");
  
  // Nombre de la hoja en A2
  var nombreHoja = hojaConsulta.getRange("A2").getValue();
  
  // Texto de fecha a buscar en B2 (ejemplo: "Thursday 27/11")
  var fechaBuscada = hojaConsulta.getRange("B2").getValue();
  
  // Obtenemos la hoja seleccionada
  var hojaSeleccionada = planilla.getSheetByName(nombreHoja);
  if (!hojaSeleccionada) {
    hojaConsulta.getRange("C2").setValue("La hoja '" + nombreHoja + "' no existe.");
    return;
  }
  
  // Definimos el rango completo donde buscar (A1:I90)
  var rangoDatos = hojaSeleccionada.getRange("A1:I90").getValues();
  var valorEncontrado = "";
  
  // Recorremos todas las filas y columnas
  for (var fila = 0; fila < rangoDatos.length; fila++) {
    for (var col = 0; col < rangoDatos[fila].length; col++) {
      var textoCelda = rangoDatos[fila][col];
      if (textoCelda && textoCelda.toString().trim() === fechaBuscada.toString().trim()) {
        // Si encontramos la fecha, tomamos el valor de la celda justo debajo
        if (fila + 1 < rangoDatos.length) {
          valorEncontrado = rangoDatos[fila + 1][col];
        }
        break;
      }
    }
    if (valorEncontrado !== "") break;
  }
  
  if (valorEncontrado === "") {
    hojaConsulta.getRange("C2").setValue("No se encontró la fecha " + fechaBuscada);
  } else {
    hojaConsulta.getRange("C2").setValue(valorEncontrado);
  }
}
