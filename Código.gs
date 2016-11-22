function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp. -> Creamos el menu con submenu
  ui.createMenu('Actualizar datos')
      .addItem('Asignar tarea', 'funcionFormularioIncidencia')
      .addItem('Limpiar hoja1', 'limpiarSheet')
      .addItem('Limpiar hoja2', 'limpiarSheet2')
      .addToUi();
}


function funcionFormularioIncidencia() {
  
  //la variable sheet selecciona la hoja actual. [0] --> Indica que la hoja que vamos a usar es la primera.
  //la variable fila selecciona la última fila de la hoja actual.
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var fila = sheet.getLastRow();
  
  //Nuestra variable "tarea" almacena el error del usuario para poder mandarselo y asignarlo a un técnico.
  //Seleccionamos la columna "H" que es en la que está almacenado el error.
 
  var tarea;
  var tarea = sheet.getRange("H" + fila).getValue();
  
  
  //Accedemos a la segunda hoja, dónde se encuentran nuestros técnicos.
  //La variable filaTecnicos selecciona la última fila de la hoja sheetTecnicos.
  
  var sheetTecnicos = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  var filaTecnicos = sheet.getLastRow();
  
  //La variable resultado la utilizamos para recorrer las filas-columnas de la tabla.
  //La variable contenido es utilizada para comprobar si la casilla está llena (mediante el .length). 
  //Cunado una casilla está llena  se pasa a la siguiente, para poder asignar las tareas al técnico que menos tiene.
  //Finalmente la última linea asigna la tarea a la casilla que le corresponde.
  
  var i = 0;
  var contenido;
  var resultado = 1;
  var letra;

  while (resultado > 0) {
        contenido = sheetTecnicos.getRange("A" + (filaTecnicos + i)).getValue();
        resultado = contenido.length;
        letra = "A";
        if (resultado > 0) {
            contenido = sheetTecnicos.getRange("B" + (filaTecnicos + i)).getValue();
            var resultado = contenido.length;
            letra = "B";
        }
        if (resultado > 0) {
            contenido = sheetTecnicos.getRange("C" + (filaTecnicos + i)).getValue();
            var resultado = contenido.length;
            letra = "C";

        }
        if (resultado > 0) {
            i++;
        }
    }

    sheetTecnicos.getRange(letra + (filaTecnicos + i)).setValue(tarea);
    
  }
  


