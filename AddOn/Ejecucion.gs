function onInstall(e) {
  var lang= Session.getActiveUserLocale(); 
  
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

function onOpen(e) {
  var lang= Session.getActiveUserLocale();
  
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

function getStarted(){
  var typeAccount= checkAccount();
  if (typeAccount == null){
    confirmation("Hay un problema", "Para poder utilizar este Add-on necesitas tener una cuenta en Google Apps.");
    return;
  }
  
  var email= Session.getActiveUser().getEmail();
  var dominio= email.substring(email.lastIndexOf("@"),email.length);
  
  var userProperties= PropertiesService.getUserProperties();
  userProperties.setProperty("unidadOrg", "/"); 
  userProperties.setProperty("dominio", dominio) 
  
  var lang= Session.getActiveUserLocale();
  
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
      .setTitle(valores[0]);
  
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function initialize(){
  var lang= Session.getActiveUserLocale();
  
  var valores= [
    "Nombre de usuario",
    "Contraseña",
    "Nombre",
    "Apellidos",
    "Correo de contacto",
    "Centro educativo",
    "Grupo",
    "Nombre de tutor 1",
    "Apellidos de tutor 1",
    "Nombre de tutor 2",
    "Apellidos de tutor 2",
    "Exito (No completar)"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  try{
    var hoja= SpreadsheetApp.getActiveSheet();
    hoja.setName("Lazy Secretary");
    hoja.clear();
    hoja.setFrozenRows(1);
    
    hoja.getRange('A1').setBackground("#CEE3F6").setValue("  "+valores[0]+" *  "); 
    hoja.getRange('B1').setBackground("#CEE3F6").setValue("  "+valores[1]+" *  "); 
    hoja.getRange('C1').setBackground("#CEE3F6").setValue("  "+valores[2]+" *  ");
    hoja.getRange('D1').setBackground("#CEE3F6").setValue("  "+valores[3]+" *  ");
    hoja.getRange('E1').setBackground("#CEE3F6").setValue("  "+valores[4]+"  ");
    hoja.getRange('F1').setValue("  "+valores[5]+"  ");
    hoja.getRange('G1').setValue("  "+valores[6]+"  ");
    hoja.getRange('H1').setValue("  "+valores[7]+"  ");
    hoja.getRange('I1').setValue("  "+valores[8]+"  ");
    hoja.getRange('J1').setValue("  "+valores[9]+"  ");
    hoja.getRange('K1').setValue("  "+valores[10]+"  ");
    hoja.getRange('L1').setValue("  "+valores[11]+"  ");
    
    for (var i=1; i<=12; i++){
      hoja.autoResizeColumn(i);
    }
    
    return true;
    
  }catch(e){
    return false;
  }
}

function createUsers(permitirEmail, cuotaMinima){
  var lang= Session.getActiveUserLocale();
  
  var valores= [
    "Los campos con '*' no pueden estar vacíos",
    "Usuario no creado: introduce un correo electrónico válido para enviar el correo de bienvenida",
    "Usuario no creado: has agotado todos los correos electrónicos que puedes enviar en un día",
    "Usuario no creado: has decidido mantener algunos mensajes de correo y ya has alcanzado el mínimo",
    "El usuario se ha creado, correo electrónico enviado",
    "El usuario se ha creado, pero has agotado todos los correos electrónicos que podías enviar en el día",
    "El usuario se ha creado, pero el correo electrónico del campo 'Correo de contacto' no es válido",
    "El usuario se ha creado. Configurado para no enviar correo electrónico de bienvenida"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var ui = SpreadsheetApp.getUi(); 
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var data = getData_();
  if (data == null){ return null; }
  
  var contador= 1;
  var usuariosCreados= 0;
  
  for (var i = 0; i < data.length; i++){
    contador ++;
    var celdaExito= sheet.getRange(contador,12);
    
    if (!checkData(data[i])){
      celdaExito.setValue(valores[0]);
      celdaExito.setBackground("#F5BCA9");
      continue;
    };
    
    var cuota= checkDailyQuota();
    
    if (permitirEmail && cuotaMinima != -1){ 
      if (!checkEmail(data[i])){ 
        celdaExito.setValue(valores[1]); 
        celdaExito.setBackground("#F5BCA9");
        continue;
      }
      
      if (cuota == 0){ 
        celdaExito.setValue(valores[2]);
        celdaExito.setBackground("#F5BCA9");
        continue;
      }
      
      if (cuota <= cuotaMinima){ 
        celdaExito.setValue(valores[3]);
        celdaExito.setBackground("#F5BCA9");
        continue;
      }
      
    }
    
    var alumno = data[i],
        usuario = alumno[0],
        pass = alumno[1],
        nombre = alumno[2],
        apellidos = alumno[3],
        correoContacto = alumno[4],
        centro = alumno[5],
        grupo = alumno[6],
        nombreTutor1 = alumno[7],
        apellidoTutor1 = alumno[8],
        nombreTutor2 = alumno[9],
        apellidoTutor2= alumno[10],
        creado= alumno[11];
    
    var userProperties= PropertiesService.getUserProperties();
    var unidadOrg= userProperties.getProperty("unidadOrg");
    var dominio= userProperties.getProperty("dominio");
    
    var exito= addUser(usuario+dominio, nombre, apellidos, pass, unidadOrg);

    if (exito == "exito"){
      if (permitirEmail){
        if (checkEmail(data[i])){
          if (cuota != 0){
            sendEmail(correoContacto, usuario+dominio, pass, nombre);
            celdaExito.setValue(valores[4]); 
            celdaExito.setBackground("#D8F6CE");
          }else{
            celdaExito.setValue(valores[5]); 
            celdaExito.setBackground("#F5DA81");
          }
          
        }else{
          celdaExito.setValue(valores[6]); 
          celdaExito.setBackground("#F5DA81");
        }
      }else{
        celdaExito.setValue(valores[7]); 
        celdaExito.setBackground("#D8F6CE");
      }
      
      usuariosCreados++;
    }else{
      celdaExito.setValue(exito);
      celdaExito.setBackground("#F5BCA9");
    }
  }
  
  return usuariosCreados;
}

function chooseUO(){
  var lang= Session.getActiveUserLocale();
  
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
      .setHeight(200)
  
  SpreadsheetApp.getUi().showModalDialog(html, valores[0]);
}

function getData_(){
  try{
    var sheet = SpreadsheetApp.getActiveSheet();
    var startRow = 2;
    var numRows = sheet.getLastRow() - 1;
    var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
    var values = data.getValues();
    
    return values;
    
  }catch (e){
    confirmation("Tenemos un problema...", "La hoja de cálculo no puede estar vacía.");
    return null;
  }
}
