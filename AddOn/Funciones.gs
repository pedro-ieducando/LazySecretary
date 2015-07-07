// Cuando se abre el documento se añaden las opciones del complemento al menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Inicializar hoja', 'initialize')
    .addItem('Elegir Unidad Org.', 'chooseUO')
    .addItem('Crear usuarios', 'createUsers')
    .addToUi();
}

// Inicializa la hoja insertando en la primera fila una serie de valores para
// indicar al usuario como debe rellenar la hoja
function initialize(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').setBackground("#CEE3F6").setValue('  NOMBRE DE USUARIO *  '); // Nombre de usuario en el dominio
  spreadsheet.getRange('B1').setBackground("#CEE3F6").setValue('  CONTRASEÑA *  '); // Contraseña que usará el usuario
  spreadsheet.getRange('C1').setBackground("#CEE3F6").setValue('  NOMBRE *  ');
  spreadsheet.getRange('D1').setBackground("#CEE3F6").setValue('  APELLIDOS *  ');
  spreadsheet.getRange('E1').setBackground("#CEE3F6").setValue('  DOMINIO *  '); // Para saber en que dominio darlo de alta
  spreadsheet.getRange('F1').setBackground("#CEE3F6").setValue('  UNIDAD ORGANIZATIVA *  '); // Unidad organizativa a la que pertenecerá
  spreadsheet.getRange('G1').setValue('  CENTRO  ');
  spreadsheet.getRange('H1').setValue('  GRUPO  ');
  spreadsheet.getRange('I1').setValue('  NOMBRE TUTOR 1  ');
  spreadsheet.getRange('J1').setValue('  APELLIDOS TUTOR 1  ');
  spreadsheet.getRange('K1').setValue('  NOMBRE TUTOR 2  ');
  spreadsheet.getRange('L1').setValue('  APELLIDOS TUTOR 2  ');
  spreadsheet.getRange('M1').setBackground("#CEE3F6").setValue('  CORREO CONTACTO *  ');
  spreadsheet.getRange('N1').setValue('  EXITO (no completar)  ');
  
  for (var i=1; i<=14; i++){
    spreadsheet.autoResizeColumn(i); // Con esto se auto dimensionan las celdas con el contenido que posean
  }
}

// Crea usuarios en el dominio a partir de los datos de la hoja
function createUsers(){
  var ui = SpreadsheetApp.getUi(); // Recogemos la interfaz
  var asked= false; // Informa de si se le ha preguntado al usuario el enviar correos hasta acabar con la cuota diaria
  var permitirEnvio= true; // Indicará si se va a permitir el envio de correos electronicos
  var sheet = SpreadsheetApp.getActiveSheet(); // Hoja actual
  
  var data = getData_(); // Recoge los datos de la hoja de calculo
  if (data == null){ return; }
  
  var contador= 1; // para saber la fila por donde vamos
  
  for (var i = 0; i < data.length; i++){
    contador ++;
    var celdaExito= sheet.getRange(contador,14); // Celda que indica el exito de la operacion actual
    
    // Comprobamos si la fila actual tiene alguno de los campos imprescindibles vacío
    if (!checkData(data[i])){
      celdaExito.setValue("Los campos con '*' no pueden estar vacíos");
      celdaExito.setBackground("#F5BCA9");
      continue;
    };
    
    // Comprobamos que el mail al que se envian las credenciales está bien formado
    if (!checkEmail(data[i])){
      celdaExito.setValue("Introduce un correo electrónico válido para enviar las credenciales!");
      celdaExito.setBackground("#F5BCA9");
      continue;
    }
    
    // Comprueba la cuota diaria restante
    var cuota= checkDailyQuota();
    Logger.log(cuota);
    if (cuota <=100 && asked == false){
      asked = true; // Una vez aqui, ya no se le volvera a preguntar de nuevo
      
      var respuesta= askYesNo(ui, "Te quedan "+cuota+" mensajes de tu cuota diaria de envíos. ¿Quieres seguir enviando correos a los usuarios?");
      
      if (respuesta == false){
        permitirEnvio= false;
      }
    }
    
    // Si no se permite seguir con el envio o se supera la cuota diaria pasa al siguiente registro
    if (!permitirEnvio || cuota == 0){
      celdaExito.setValue("Usuario no creado. La cuota diaria de envío de correo se ha \nsuperado o se ha denegado por el propio usuario.");
      celdaExito.setBackground("#F5BCA9");
      continue;
    }
    
    var alumno = data[i],
        usuario = alumno[0],
        pass = alumno[1],
        nombre = alumno[2],
        apellidos = alumno[3],
        dominio = alumno[4],
        unidadOrg = alumno[5],
        centro = alumno[6],
        grupo = alumno[7],
        nombreTutor1 = alumno[8],
        apellidoTutor1 = alumno[9],
        nombreTutor2 = alumno[10],
        apellidoTutor2= alumno[11],
        correoContacto = alumno[12],
        creado= alumno[14];
    
    // Si el alumno ya ha sido creado con anterioridad pasa al siguiente registro
    if (creado == "Alumno creado"){
      continue;
    }
    
    var exito= addUser(usuario+"@"+dominio, nombre, apellidos, pass, unidadOrg);
    if (exito == "exito"){
      sendEmail(correoContacto, usuario+"@"+dominio, pass, nombre);
      celdaExito.setValue("Alumno creado");
      celdaExito.setBackground("#D8F6CE");
    }else{
      celdaExito.setValue(exito);
      celdaExito.setBackground("#F5BCA9");
    }
  } 
}

// Abre un cuadro de dialogo que permite seleccionar una unidad organizativa
function chooseUO(){
  var html = HtmlService.createHtmlOutputFromFile('chooseUO')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(750)
      .setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecciona Unidad Organizativa');
}

// Recoge los datos de la hoja de calculo, a partir de la segunda fila
function getData_(){
  var ui = SpreadsheetApp.getUi(); // Recogemos la interfaz
  
  try{
    var sheet = SpreadsheetApp.getActiveSheet();
    var startRow = 2;
    var numRows = sheet.getLastRow() - 1;
    var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
    var values = data.getValues();
    
    return values;
    
  }catch (e){
    ui.alert(
      'La hoja de cálculo esta vacía!',
      ui.ButtonSet.OK);
    return null;
  }
}
