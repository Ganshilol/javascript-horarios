function crearIndice() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hojaIndice = planilla.getSheetByName("consulta");
  var hojas = planilla.getSheets();
  var total = hojas.length;

  // Mostrar el total de hojas en el log
  Logger.log("El total de hojas de este archivo es: " + total);

  var arrNombreHojas = [];

  // Recorremos solo desde la hoja 6 hasta la 94
  // Ojo: los índices en Apps Script empiezan en 0, así que la hoja 6 es índice 5
  for (var i = 6; i < 94 && i < total; i++) {
    arrNombreHojas.push([hojas[i].getName()]);
    Logger.log(hojas[i].getName());
  }

  Logger.log(arrNombreHojas);

  // Escribimos en la hoja "consulta" solo esas hojas
  hojaIndice.getRange(3, 1, arrNombreHojas.length, 1).setValues(arrNombreHojas);
}
