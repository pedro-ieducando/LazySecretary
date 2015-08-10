// Cuando se instala el documento se añaden las opciones del complemento al menu
function onInstall(e) {
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores=[
    "Crear usuarios"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem(valores[0], 'getStarted')
    .addToUi();
  
  initialize();
  getStarted();
}

// Cuando se abre el documento se añaden las opciones del complemento al menu
function onOpen(e) {
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores=[
    "Crear usuarios"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem(valores[0], 'getStarted')
    .addToUi();
}


// Inicializa la hoja insertando en la primera fila una serie de valores para
// indicar al usuario como debe rellenar la hoja
function initialize(){
  
  // Comprueba que tiene una cuenta de Google Apps
  var typeAccount= checkAccount();
  if (typeAccount == false){
    confirmation("Hay un problema", "Para poder utilizar este Add-on necesitas tener una cuenta en Google Apps.");
    return;
  }
  
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Nombre de usuario",
    "Contraseña",
    "Nombre",
    "Apellidos",
    "Dominio",
    "Unidad Organizativa",
    "Centro educativo",
    "Grupo",
    "Nombre de tutor 1",
    "Apellidos de tutor 1",
    "Nombre de tutor 2",
    "Apellidos de tutor 2",
    "Correo de contacto",
    "Exito (No completar)"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var spreadsheet = SpreadsheetApp.getActive(); // Recogemos la hoja actual
  
  try{
    var nuevaHoja= spreadsheet.insertSheet("Lazy Secretary"); // Y creamos una nueva para estar seguros de no sobre escribir datos
    
    nuevaHoja.getRange('A1').setBackground("#CEE3F6").setValue("  "+valores[0]+" *  "); // Nombre de usuario en el dominio
    nuevaHoja.getRange('B1').setBackground("#CEE3F6").setValue("  "+valores[1]+" *  "); // Contraseña que usará el usuario
    nuevaHoja.getRange('C1').setBackground("#CEE3F6").setValue("  "+valores[2]+" *  ");
    nuevaHoja.getRange('D1').setBackground("#CEE3F6").setValue("  "+valores[3]+" *  ");
    nuevaHoja.getRange('E1').setBackground("#CEE3F6").setValue("  "+valores[4]+" *  "); // Para saber en que dominio darlo de alta
    nuevaHoja.getRange('F1').setBackground("#CEE3F6").setValue("  "+valores[5]+" *  "); // Unidad organizativa a la que pertenecerá
    nuevaHoja.getRange('G1').setValue("  "+valores[6]+"  ");
    nuevaHoja.getRange('H1').setValue("  "+valores[7]+"  ");
    nuevaHoja.getRange('I1').setValue("  "+valores[8]+"  ");
    nuevaHoja.getRange('J1').setValue("  "+valores[9]+"  ");
    nuevaHoja.getRange('K1').setValue("  "+valores[10]+"  ");
    nuevaHoja.getRange('L1').setValue("  "+valores[11]+"  ");
    nuevaHoja.getRange('M1').setBackground("#CEE3F6").setValue("  "+valores[12]+" *  ");
    nuevaHoja.getRange('N1').setValue("  "+valores[13]+"  ");
    
    for (var i=1; i<=14; i++){
      nuevaHoja.autoResizeColumn(i); // Con esto se auto dimensionan las celdas con el contenido que posean
    }
  }catch(e){
    confirmation("Error al inicializar", "Ya tienes una hoja con el nombre 'Lazy Secretary'.")
    return;
  }
}

// Abre una sideBar para guiar al usuario en el proceso de creación de usuarios 
function getStarted(){
  // Comprueba que tiene una cuenta de Google Apps
  var typeAccount= checkAccount();
  if (typeAccount == false){
    confirmation("Hay un problema", "Para poder utilizar este Add-on necesitas tener una cuenta en Google Apps.");
    return;
  }
  
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Crear usuarios"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var html = HtmlService.createHtmlOutputFromFile('sideBar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle(valores[0])
      .setWidth(200);
  
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

// Crea usuarios en el dominio a partir de los datos de la hoja
function createUsers(permitirEmail, cuotaMinima){
  //enviar mail es la opcion que te permite decidir si quieres enviar correos electronicos a los nuevos usuarios que se vayan a crear
  
  // Comprueba que tiene una cuenta de Google Apps
  var typeAccount= checkAccount();
  if (typeAccount == false){
    confirmation("Hay un problema", "Para poder utilizar este Add-on necesitas tener una cuenta en Google Apps.");
    return;
  }
  
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Los campos con '*' no pueden estar vacíos",
    "Introduce un correo electrónico válido para enviar el correo de bienvenida",
    "No se creará el usuario por no poder enviar el correo de bienvenida. Se ha llegado al mínimo establecido.",
    "Has decidido no enviar más correos electrónicos. ¿Quieres crear los usuarios de todas formas? (no se les enviará correo de bienvenida)",
    "Se ha acabado tu cuota diaria de envío de correo electrónico. ¿Quieres crear los usuarios de todas formas? (no se les enviará correo de bienvenida)",
    "Usuario no creado por decisión del Administrador",
    "Usuario creado"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var ui = SpreadsheetApp.getUi(); // Recogemos la interfaz
  var sheet = SpreadsheetApp.getActiveSheet(); // Hoja actual
  
  var data = getData_(); // Recoge los datos de la hoja de calculo
  if (data == null){ return; }
  
  var contador= 1; // para saber la fila por donde vamos
  var usuariosCreados= 0; // para llevar un control de los usuarios creados
  
  for (var i = 0; i < data.length; i++){
    contador ++;
    var celdaExito= sheet.getRange(contador,14); // Celda que indica el exito de la operacion actual
    
    // Comprobamos si la fila actual tiene alguno de los campos imprescindibles vacío
    if (!checkData(data[i])){
      celdaExito.setValue(valores[0]);
      celdaExito.setBackground("#F5BCA9");
      continue;
    };
    
    // Comprueba la cuota diaria restante
    var cuota= checkDailyQuota();
    
    if (permitirEmail && cuotaMinima != -1){ 
      if (!checkEmail(data[i])){
        celdaExito.setValue(valores[1]); // Si el correo no es valido pasa al siguiente
        celdaExito.setBackground("#F5BCA9");
        continue;
      }
      
      if (cuota <= cuotaMinima){ // Si se supera la cuota minima establecida termina
        celdaExito.setValue(valores[2]);
        celdaExito.setBackground("#F5BCA9");
        continue;
      }
    }
    
    
    
    // En desarrollo
    // Se debe tener en cuenta que si no se establece cuota minima, lo que hacer si el correo esta mal formado
    
    
    
    
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
        creado= alumno[13];
    
    // Si el alumno ya ha sido creado con anterioridad pasa al siguiente registro
    if (creado == valores[6]){
      continue;
    }
    
    // Realizando pruebas
    //var exito= addUser(usuario+"@"+dominio, nombre, apellidos, pass, unidadOrg);
    
    var exito="exito";
    if (exito == "exito"){
      if (permitirEmail){
        sendEmail(correoContacto, usuario+"@"+dominio, pass, nombre);
      }
      
      celdaExito.setValue(valores[6]);
      celdaExito.setBackground("#D8F6CE");
      usuariosCreados++;
    }else{
      celdaExito.setValue(exito);
      celdaExito.setBackground("#F5BCA9");
    }
  }
  
  confirmation("Trabajo finalizado", "Ha terminado la ejecución del programa. Se crearon "+usuariosCreados+" usuarios.");
  //return usuariosCreados;
}

// Abre un cuadro de dialogo que permite seleccionar una unidad organizativa
function chooseUO(){
  
  // Comprueba que tiene una cuenta de Google Apps
  var typeAccount= checkAccount();
  if (typeAccount == false){
    confirmation("Hay un problema", "Para poder utilizar este Add-on necesitas tener una cuenta en Google Apps.");
    return;
  }
  
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Seleccionar Unidad Organizativa"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  
  var html = HtmlService.createHtmlOutputFromFile('chooseUO')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(600)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, valores[0]);
}


// Recoge los datos de la hoja de calculo, a partir de la segunda fila
function getData_(){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "La hoja de cálculo esta vacia"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
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
      valores[0],
      ui.ButtonSet.OK);
    return null;
  }
}
