// Comprueba si todos los campos obligatorios estan rellenados. En caso de que
// haya alguno sin inicializar, devolvera false
function checkData(alumno){  
  for (var i=0; i<6; i++){
    if (alumno[i] == ""){
      return false;
    }
  }
  
  if(alumno[12] == ""){
    return false;
  }
  
  return true;
}


// Comprueba si la direccion de correo a la que se envian las credenciales está bien formado
function checkEmail(alumno){
  var email = alumno[12],
      destino, // Nombre al que va dirigido el correo
      dominio="", // Lo que va despues de la @
      arroba, // Posicion de la @ en la direccion de correo
      punto, // Posicion del ultimo punto del dominio
      beforePunto, // Lo que hay entre la @ y el ultimo punto
      afterPunto; // Lo que hay despues del ultimo punto
  
  email= email.trim(); // Quitamos los espacios al final y al principio por si acaso
  
  // Comprobamos que no haya espacios por medio
  for (var i=0; i<email.length; i++){
    var c= email[i];
    if (c == " "){
      return false;
    }
  }
  
  // Que haya una @ y solo una
  if (email.indexOf("@") != -1 && (email.indexOf("@") == email.lastIndexOf("@"))){
    arroba= email.indexOf("@");
    dominio= email.substring(arroba+1, email.length); 
    
    // Que tenga una parte de dominio definida
    if (dominio.length == 0){
      return false;
    }
  }else{
    return false;
  }
  
  punto= dominio.lastIndexOf(".");
  if (punto != -1){
    beforePunto= dominio.substring(0, punto);
    afterPunto= dominio.substring(punto+1, dominio.length);
    
    if (beforePunto.length == 0 || afterPunto.length == 0){
      return false
    }
  }else{
    return false;
  }
  
  return true;
}

// Devuelve la cuota diaria de mails que queda
function checkDailyQuota(){
  return MailApp.getRemainingDailyQuota();
}

// Muestra un mensaje de confirmacion preguntando el mensaje que se le pasa por parametros
function askYesNo(ui, mensaje){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Por favor, confirma",
    mensaje
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var result = ui.alert(
     valores[0],
     valores[1],
     ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    return true;
  } else {
    return false;
  }
}

// Muestra un mensaje informativo que se le pasa por parametros
function confirmation(title, msg){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    title,
    msg
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var ui = SpreadsheetApp.getUi(); // Recogemos la interfaz
  var result = ui.alert(
     valores[0],
     valores[1],
     ui.ButtonSet.OK);
}

function prueba(){
  var ui= SpreadsheetApp.getUi();
  //var ui.ButtonSet.
}


// Inserta un usuario en el dominio. Controla varias excepciones
function addUser(email, nombre, apellidos, pass, orgUnit){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "El usuario ya existe",
    "Necesitas privilegios para añadir usuarios a tu dominio",
    "No se encuentra la unidad organizativa",
    "El dominio introducido no es válido",
    "El correo electrónico del nuevo usuario no es válido",
    "La contraseña no es válida"
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var user = {
    primaryEmail: email,
    name: {
      givenName: nombre,
      familyName: apellidos
    },
    password: pass,
    orgUnitPath: orgUnit
  };
  try{
    user = AdminDirectory.Users.insert(user);
    return "exito";
  }catch(e){
    // Tratamiento de errores
    if (e == "Exception: Entity already exists."){
      return valores[0];
    }
    if (e == 'ReferenceError: "AdminDirectory" no está definido.'){
      return valores[1];
    }
    if (e == "Exception: Invalid Input: INVALID_OU_ID"){
      return valores[2];
    }
    if (e == "Exception: Resource Not Found: domain"){
      return valores[3];
    }
    if (e == "Exception: Invalid Input: primary_user_email"){
      return valores[4];
    }
    if (e == "Exception: Invalid Password"){
      return valores[5];
    }
    
    return "Error: "+e;
  }
}


// Envia un mensaje al destinatario que se le pasa por parametros
function sendEmail(destinatario, usuario, pass, nombre){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "Credenciales de acceso",
    "Hola",
    "Se ha creado una cuenta corporativa para ti en el entorno de Google Apps. En este entorno podrás disfrutar nuestro correo corporativo y muchos más beneficios.. Tus credenciales de acceso son:",
    "Nombre de usuario: ",
    "Clave de acceso: ",
    "Para acceder tienes que hacer login en Google con tu usuario y la contraseña. La primera vez que entres se te solicitará que cambies la contraseña y que establezcas una de tu elección.",
    "Para acceder debes dirigirte a: ",
    "Con tu nueva cuenta corporativa podrás realizar muchísimas acciones y además tendrás fantásticas ventajas: ",
    "Espacio ilimitado en la nube: podrás sincronizar con tus dispositivos archivos sin límite de tamaño ni problemas de espacio (como Dropbox, pero de forma ilimitada).",
    "Sin publicidad: no se muestra publicidad de Google ni en las búsquedas ni en correo.",
    "Sin Spam ni virus: Google filtra todo lo que entra en tu bandeja de entrada.",
    "En el entorno de Google Apps encontrarás las siguientes aplicaciones integradas y que podrás disfrutar desde tu cuenta corporativa:",
    "Comunicación",
    "Correo electrónico corporativo y contactos integrados en Gmail.",
    "Conéctate con las personas que quieras mediante llamadas de voz, chat de texto o vídeo de alta definición. Puedes ahorrar tiempo y dinero en viajes, sin renunciar a ninguna de las ventajas del contacto cara a cara.",
    "Dedica menos tiempo a la planificación y más al trabajo con los calendarios, que se pueden compartir y se integran perfectamente con Gmail, Drive, Contactos, Sites y Hangouts, para que puedas saber en todo momento cuál es el próximo evento.",
    "Red social en el entorno corporativo. Podrás compartir enlaces, videos, comentarios y darte de alta en grupos afines.",
    "Almacenamiento",
    "Mantén todo tu trabajo en un lugar seguro con el almacenamiento de archivos online. Accede a tu trabajo cuando lo necesites, desde el portátil, el tablet o el teléfono móvil.",
    "Colaboración",
    "Crea y edita documentos de texto directamente en tu navegador sin necesidad de software específico. Pueden trabajar varias personas al mismo tiempo en un archivo: todos los cambios se guardan automáticamente.",
    "Crea hojas de cálculo directamente en tu navegador sin necesidad de software específico. Puedes utilizarlas para todo tipo de contenido, desde sencillas listas de tareas hasta análisis de datos con gráficos, filtros y tablas dinámicas.",
    "Crea formularios personalizados para encuestas y cuestionarios sin ningún cargo adicional. Recopila toda la información en una hoja de cálculo y analiza los datos directamente en Hojas de cálculo de Google.",
    "Crea y edita elegantes presentaciones en tu navegador sin necesidad de software específico. Pueden trabajar varias personas al mismo tiempo; de esta forma, todos tienen siempre la versión más reciente.",
    "Crea un sitio de proyectos para tu equipo con nuestra aplicación para el diseño de sitios web. Y todo sin escribir ni una sola línea de código.",
    "Y muchas herramientas más..."
  ];
  
  if (lang != "es"){
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i], lang);
    }
  }
  
  var dominio= usuario.substring(usuario.lastIndexOf("@")+1, usuario.length);
  
  MailApp.sendEmail({
     to: destinatario,
     subject: valores[0],
    htmlBody: 
    '<div id="cuerpo" style="text-align: justify; font-family: Arial, Helvetica, sans-serif; color:#2E2E2E">'+valores[1]+', '+(nombre.charAt(0).toUpperCase() + nombre.slice(1))+'!<br><br>'
      +valores[2]+'<br><br>'
      +valores[3]+"<b> "+usuario+"</b><br>"
      +valores[4]+"<b> "+pass+"</b> <br><br>"
      +valores[5]+'<br><br>'
      +valores[6]+' <a href="https://accounts.google.com">https://accounts.google.com</a><br><br>'

      +valores[7]
      +'<ul><li>'+valores[8]+'</li>'
      +'<li>'+valores[9]+'</li>'
      +'<li>'+valores[10]+'</li></ul>'
      +valores[11]+'<br><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[12]+'</div>'

      +'<!-- Gmail --><img src="https://imagizer.imageshack.us/v2/64x46q90/r/537/Oqyt4a.png" /><br /><a style="font-size:14pt;" href="http://mail.google.com">Gmail</a>'
      +'<div style="color:#656565;">'+valores[13]+'.</div><br><br>'

      +'<!-- Talk / Hangouts --><img src="https://imagizer.imageshack.us/v2/64x64q90/r/673/X4UitP.png" /><br /><a style="font-size:14pt;" href="http://www.google.es/hangouts/">Talk / Hangouts</a>'
      +'<div style="color:#656565;">'+valores[14]+'</div><br><br>'

      +'<!-- Calendar --><img src="https://imagizer.imageshack.us/v2/64x54q90/r/540/Nd1Pcp.png" /><br /><a style="font-size:14pt;" href="http://calendar.google.com">Calendar</a>'
      +'<div style="color:#656565;">'+valores[15]+'</div><br><br>'

      +'<!-- Google+ --><img src="https://imagizer.imageshack.us/v2/64x64q90/r/538/kj10iP.png" /><br /><a style="font-size:14pt;" href="http://plus.google.com">Google+</a>'
      +'<div style="color:#656565;">'+valores[16]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[17]+'</div>'

      +'<!-- Drive --><img src="https://imagizer.imageshack.us/v2/64x53q90/r/537/zNdrGw.png" /><br /><a style="font-size:14pt;" href="http://drive.google.com">Drive</a>'
      +'<div style="color:#656565;">'+valores[18]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[19]+'</div>'

      +'<!-- Docs --><img src="https://imagizer.imageshack.us/v2/52x64q90/r/661/sQ8R9g.png" /><br /><a style="font-size:14pt;" href="http://docs.google.com">Docs</a>'
      +'<div style="color:#656565;">'+valores[20]+'</div><br><br>'

      +'<!-- Sheets --><img src="https://imagizer.imageshack.us/v2/50x64q90/r/910/QUzX4t.png" /><br /><a style="font-size:14pt;" href="http://sheets.google.com">Sheets</a>'
      +'<div style="color:#656565;">'+valores[21]+'</div><br><br>'

      +'<!-- Forms --><img src="https://imagizer.imageshack.us/v2/53x64q90/r/540/KJnYH6.png" /><br /><a style="font-size:14pt;" href="http://forms.google.com">Forms</a>'
      +'<div style="color:#656565;">'+valores[22]+'</div><br><br>'

      +'<!-- Slides --><img src="https://imagizer.imageshack.us/v2/52x64q90/r/537/cqAZKA.png" /><br /><a style="font-size:14pt;" href="http://slides.google.com">Slides</a>'
      +'<div style="color:#656565;">'+valores[23]+'</div><br><br>'

      +'<!-- Sites --><img src="https://imagizer.imageshack.us/v2/64x56q90/r/661/blhVle.png" /><br /><a style="font-size:14pt;" href="http://sites.google.com">Sites</a>'
      +'<div style="color:#656565;">'+valores[24]+'</div><br><br>'

    +'<a href="http://www.google.es/about/products/">'+valores[25]+'</a>'
   });
}


// Recoge las unidades organizativas del dominio
function cargarUO(){
  try{
    var mailUser= Session.getActiveUser().getEmail();
    var customerId= AdminDirectory.Users.get(mailUser).customerId;
    
    var orgUnits= AdminDirectory.Orgunits.list(customerId, {type:"all"}).organizationUnits;
    
    orgUnits.sort(orderUO);
    
    return orgUnits;
  }catch(e){
    return null;
  }
  
}

// Ordena las unidades organizativas por ruta
function orderUO(a,b) {
  if (a.orgUnitPath.toLowerCase() < b.orgUnitPath.toLowerCase()){ return -1; }
  if (a.orgUnitPath.toLowerCase() > b.orgUnitPath.toLowerCase()){ return 1; }
  return 0;
}

// Cuando una unidad organizativa es seleccionada, se modifica la celda correspondiente para todos y cada uno de los usuarios
function UOSelected(selected){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = getData_(); // Recoge los datos de la hoja de calculo
  if (data == null){ return; }
  
  var contador= 1; // para saber la fila por donde vamos
  
  for (var i = 0; i < data.length; i++){
    contador ++;
    var celdaUO= sheet.getRange(contador,6);
    celdaUO.setValue(selected); // Cambia el valor de la celda por la unidad organizativa seleccionada
  }
}

// traduce cadenas de texto para compatibilizar el complemento con otros idiomas
function translate(texto, idioma){
  var traducido= LanguageApp.translate(texto, 'es', idioma);
  return traducido;
}

// comprueba que tipo de cuenta tiene el usuario (Google account o Google Apps)
function checkAccount(){
  var mailUser= Session.getActiveUser().getEmail();
  var dominio= mailUser.substring(mailUser.lastIndexOf("@"));
  
  if (dominio == "@gmail.com"){
    return false;
  }else{
    return true;
  }
}
