// Nombre de la hoja : Enviar Correo Aviso

function enviarCorreo() {
  
  /*
  AquÍ decimos que sea la hoja activa
  Es un “bound” script (un código adjunto a un documento)
  
  */
  var sheet = SpreadsheetApp.getActive()
  
  var rows =sheet.getDataRange(); /*traemos todas las filas que tiene la hoja*/
  var numRows = rows.getNumRows() -1; /*traemos el número de la última fila  y le restamos 1*/
  var values = rows.getValues(); /*traer los valores que hay en las filas*/
  
  /*Hacemos un bucle que recorra fila por fila desde la 1 hasta la última fila y traiga los datos de las columnas*/
  for(var i=1; i<=numRows; i++){
    var emailCopia= values[i][4]; // columna 4 empezando por el cero
    var nombre= values[i][1]; // columna 1 empezando por el cero
    var email= values[i][3]; // columna 3 empezando por el cero
    var territorio=values[i][0]; // columna 0 empezando por el cero
    var id =values[i][2]; // columna 2 empezando por el cero


    var archivo = DriveApp.getFileById(id); // traemos el archivo correspondiente por su id que está en la columna 2

    //creamos el mensaje en html
    var mensaje = "Estimada(o) "+ nombre +" <br><br> Buenos días, por medio de la presente reciba un cordial saludo así mismo le escribo para solicitar su apoyo para solventar las comprobaciones de combustible pendientes de los periodos: junio, julio, agosto, septiembre y octubre. <br><br> Las comprobaciones pendientes de su territorio se anexan en un archivo a este correo. En éste se desglosa: el periodo, el responsable, la placa y los monederos de los cuales aún no se tiene una comprobación correcta. Además, se menciona el error específico del archivo enviado o en su caso si éste no ha sido enviado. <br><br> Es importante mencionar que la fecha límite para la comprobación combustible de estos periodos ya venció, por tanto, se le solicita de manera urgente atienda este correo y solvente de manera correcta las comprobaciones pendientes antes de las 10 am del día lunes 6 de diciembre de 2021.<br><br> En caso contrario, se procederá al bloqueo de los monederos que no cuenten con la comprobación correcta y completa.<br><br> Dirección de Logística y Seguimiento para el Desarrollo Rural y productivo";
    

    /*Enviar Email mediante MailApp*/
     MailApp.sendEmail({
     to: email, // para
     cc:emailCopia, //con copia para
     subject: "Solicitud Urgente", //asunto
     htmlBody: mensaje, // en el mensaje podemos incluir código html
     attachments: [archivo], // adjuntos [archivo1,archivo2,archivo 3],
   })

    
  } /*aqui termina for*/
  

       
} /*aqui termina function*/ 
