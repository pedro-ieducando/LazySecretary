// Comprueba si todos los campos obligatorios estan rellenados. En caso de que
// haya alguno sin inicializar, devolvera false
function checkData(alumno){  
  for (var i=0; i<4; i++){
    if (alumno[i] == ""){
      return false;
    }
  }
  
  return true;
}


// Comprueba si la direccion de correo a la que se envian las credenciales está bien formado
function checkEmail(alumno){
  var email = alumno[4],
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


// Inserta un usuario en el dominio. Controla varias excepciones
function addUser(email, nombre, apellidos, pass, orgUnit){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  
  // Valores de las cadenas de texto
  var valores= [
    "El usuario ya existe",
    "Necesitas privilegios para añadir usuarios a tu dominio",
    "No se encuentra la unidad organizativa",
    "No es posible encontrar el dominio especificado. Tienes que introducir \n"+"el dominio principal de tu cuenta de Google Apps, o un alias de dominio  \n"+"con estado 'activo'. Por ejemplo, para crear 'judith.smith@example.com', \n"+"el dominio es 'example.com'.",
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

      +'<!-- Gmail --><img src="http://i.imgur.com/LpS8rQ9.png" /><br /><a style="font-size:14pt;" href="http://mail.google.com">Gmail</a>'
      +'<div style="color:#656565;">'+valores[13]+'.</div><br><br>'

      +'<!-- Hangouts --><img src="http://i.imgur.com/qnaPCvu.png" /><br /><a style="font-size:14pt;" href="http://www.google.es/hangouts/">Hangouts</a>'
      +'<div style="color:#656565;">'+valores[14]+'</div><br><br>'

      +'<!-- Calendar --><img src="http://i.imgur.com/PGRnI53.png" /><br /><a style="font-size:14pt;" href="http://calendar.google.com">Calendar</a>'
      +'<div style="color:#656565;">'+valores[15]+'</div><br><br>'

      +'<!-- Google+ --><img src="http://i.imgur.com/wo8joYb.png" /><br /><a style="font-size:14pt;" href="http://plus.google.com">Google+</a>'
      +'<div style="color:#656565;">'+valores[16]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[17]+'</div>'

      +'<!-- Drive --><img src="http://i.imgur.com/o9cPX2V.png" /><br /><a style="font-size:14pt;" href="http://drive.google.com">Drive</a>'
      +'<div style="color:#656565;">'+valores[18]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[19]+'</div>'

      +'<!-- Docs --><img src="http://i.imgur.com/lM48l47.png" /><br /><a style="font-size:14pt;" href="http://docs.google.com">Docs</a>'
      +'<div style="color:#656565;">'+valores[20]+'</div><br><br>'

      +'<!-- Sheets --><img src="http://i.imgur.com/uGM3hpI.png" /><br /><a style="font-size:14pt;" href="http://sheets.google.com">Sheets</a>'
      +'<div style="color:#656565;">'+valores[21]+'</div><br><br>'

      +'<!-- Forms --><img src="http://i.imgur.com/D2y8TY1.png" /><br /><a style="font-size:14pt;" href="http://forms.google.com">Forms</a>'
      +'<div style="color:#656565;">'+valores[22]+'</div><br><br>'

      +'<!-- Slides --><img src="http://i.imgur.com/OM6qd8F.png" /><br /><a style="font-size:14pt;" href="http://slides.google.com">Slides</a>'
      +'<div style="color:#656565;">'+valores[23]+'</div><br><br>'

      +'<!-- Sites --><img src="http://i.imgur.com/8j9Nkl4.png" /><br /><a style="font-size:14pt;" href="http://sites.google.com">Sites</a>'
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
  var userProperties= PropertiesService.getUserProperties();
  userProperties.setProperty("unidadOrg", selected);
}

// devuelve el idioma del usuario
function getLanguage(){
  var lang= Session.getActiveUserLocale(); // Idioma del usuario
  return lang;
}

// traduce cadenas de texto para compatibilizar el complemento con otros idiomas
function translate(texto, idioma){
  var traducido= LanguageApp.translate(texto, 'es', idioma);
  return traducido;
}

// Comprueba qué tipo de cuenta tiene el usuario. Si el id de cliente no esta definido quiere decir
// que lo que posee una cuenta personal, no una cuenta de google apps. En ese caso devolveria nulo
// para indicar al usuario que no podrá usar el complemento
function checkAccount(){
  try{
    var mailUser= Session.getActiveUser().getEmail();
    var customerId= AdminDirectory.Users.get(mailUser).customerId;
    
    return customerId;
  }catch(e){
    return null;
  }
}
