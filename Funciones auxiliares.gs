function columnToLetter(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function encontrarColumna(hoja_, cabecera) {
  var celdaCabecera = hoja_.createTextFinder(cabecera).findNext();
  if (celdaCabecera) {
    return celdaCabecera.getColumn();
  } else {
    return false;
  }
}

function primerEventoAno(fecha) {
  const ano = fecha.getFullYear();
  const mes = fecha.getMonth();
  const dia = fecha.getDate();
  const inicio = new Date(ano, 0, 1);
  const final = new Date(ano, mes, dia-1, 23, 59)
  if (calendario.getEvents(inicio, fecha).length == 0) {
    return true
  } else {
    return false
  }
}

function getDuration(range) {
  var value = range.getValue();
  // Get the date value in the spreadsheet's timezone.
  var spreadsheetTimezone = range.getSheet().getParent().getSpreadsheetTimeZone();
  var dateString = Utilities.formatDate(value, spreadsheetTimezone, 
      'EEE, d MMM yyyy HH:mm:ss');
  var date = new Date(dateString);

  // Initialize the date of the epoch.
  var epoch = new Date('Dec 30, 1899 00:00:00');

  // Calculate the number of milliseconds between the epoch and the value.
  var diff = date.getTime() - epoch.getTime();
  return diff
}

function conseguirColor(celda) {
  var color = null
  if (celda.getMergedRanges().length == 0){
    color = celda.getBackground()
  } else if (celda.getMergedRanges().length != 0) {
    color = celda.getMergedRanges()[0].getBackground()
  } else {
    throw Error()
  }
  return color
}

function getColores(color, colores) {
  if (color === colores[0]) {
    return ['#b4a7d6','#d9d2e9']
  } else if (color === colores[1]) {
    return ['#ffe599', '#fff2cc']
  } else {
    throw new Error('El color del día anterior no es válido.')
  }
}

function alternarColor(color, colores) {
  if (color === colores[0]) {
    return colores[1];
  } else if (color === colores[1]) {
    return colores[0];
  } else if (color === '#a4c2f4') {
    return colores[0]
  } else {
    throw new Error('El color no es válido.\n'+color);
  }
}

function colorAlterno(celda, colores, hoja) {
  const col = celda.getColumn();
  const fila = celda.getRow();
  var color = null;

  for (var i = fila - 1; i > 0; i--) {  // Cambié la condición del bucle para evitar error
    color = hoja.getRange(i, col).getBackground();
    if (colores.includes(color)) {
      color = alternarColor(color, colores);
      break;
    }
  }

  if (color === null) {
    throw new Error('No se encontró ninguno de los colores de la lista.');
  }

  celda.setBackground(color);
}

function dolarBlue (fecha) {
  const fechaISO = new Date(fecha.getTime() - (fecha.getTimezoneOffset()*60*1000)).toISOString().split('T')[0];
  const url = 'https://api.bluelytics.com.ar/v2/historical?day=' + fechaISO;
  const datos = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  return datos.blue.value_avg
}

function dolarBlueMes (ano, mes) {
  var total = 0;
  var pasos = 0;
  for (let dia = new Date(ano, mes, 1); dia.getMonth()==mes; dia.setDate(dia.getDate() + 1)) {
    total += dolarBlue(dia);
    pasos += 1;
  }
  return total/pasos
}

function canastaBasica(fecha) {
  const fechaISO = `${fecha.getFullYear()}-${fecha.getMonth()+1}`;
  const url = 'https://apis.datos.gob.ar/series/api/series/?'+'ids=150.1_CSTA_BARIA_0_D_26&format=json&metadata=none'+`&start_date=${fechaISO}`+`&end_date=${fechaISO}`;
  const datos = JSON.parse(UrlFetchApp.fetch(url).getContentText()).data;
  if (datos.length == 0) {
    throw `No hay datos de la canasta básica para ese mes (${fechaISO}).`
  } else {
    const canasta = datos[0][1]
    return canasta
  }
}

function getDuracionString (float, modo='horas') {
  let segundos;
  let minutos;
  let horas;
  if (modo='horas') {
    horas = Math.floor(float);
    minutos = Math.floor(float*60-horas*60);
    segundos = (float*60*60-horas*60*60-minutos*60).toFixed(3);
  } else if (modo='minutos') {
    horas = Math.floor(float/60);
    minutos = Math.floor(float-horas*60);
    segundos = (float*60-horas*60*60-minutos*60).toFixed(3);
  } else if (modo='segundos') {
    horas = Math.floor(float/60/60);
    minutos = Math.floor(float/60-horas*60);
    segundos = (float-horas*60*60-minutos*60).toFixed(3);
  } else if (modo='milisegundos') {
    horas = Math.floor(float/1000/60/60);
    minutos = Math.floor(float/1000/60-horas*60);
    segundos = (float/1000-horas*60*60-minutos*60).toFixed(3);
  } else if (modo='dias') {
    horas = Math.floor(float*24);
    minutos = Math.floor(float*24*60-horas*60);
    segundos = (float*24*60*60-horas*60*60-minutos*60).toFixed(3);
  } else if (modo='anos') {
    horas = Math.floor(float*365*24);
    minutos = Math.floor(float*365*24*60-horas*60);
    segundos = (float*365*24*60*60-horas*60*60-minutos*60).toFixed(3);
  } else {
    throw 'El modo introducido es inválido.'
  }
  return horas+':'+minutos+':'+segundos
}

function getDuracion (celda, modo='display') {
  const fecha_base = new Date('1899-12-30T04:16:48.000Z');
  const fecha = celda.getValue();
  const milisegundos = fecha.getTime() - fecha_base.getTime();
  const segundos = milisegundos/1000;
  const minutos = segundos/60;
  const horas = minutos/60;
  const dias = horas/24;
  const anos = dias/365;
  if (modo='display') {
    return getDuracionString(horas)
  } else if (modo=='segundos') {
    return segundos
  } else if (modo=='minutos') {
    return minutos
  } else if (modo=='horas') {
    return horas
  } else if (modo=='dias') {
    return dias
  } else if (modo=='anos') {
    return anos
  } else if (modo=='milisegundos') {
    return milisegundos
  } else {
    throw 'El modo introducido es inválido.'
  }
}

// Función para configurar el disparador de evento instalable
function configurarDisparador() {
  ScriptApp.newTrigger('onEventUpdated')
      .forUserCalendar('inakigova@gmail.com') // Reemplaza con la dirección de correo electrónico del usuario
      .onEventUpdated()
      .create();
}

// Función que se ejecuta cuando se actualiza un evento en el calendario
function onEventUpdated(e) {
  // Obtener el ID del calendario donde ocurrió el evento
  var calendarId = e.calendarId;
  if (calendarId=='057c4a0768093a7c3ce599e91e3dd992dbf940d5e03ad8878eb8b6742a836062@group.calendar.google.com') {
  // Realizar una sincronización incremental de los eventos del calendario
    sincronizarEventos(calendarId);
  }
}

// Función para realizar una sincronización incremental de los eventos del calendario
function sincronizarEventos(calendarId) {
  // Obtener el token de sincronización almacenado
  var storedSyncToken = PropertiesService.getUserProperties().getProperty(calendarId + '_syncToken');

  // Realizar la solicitud de sincronización incremental
  var syncParams = {
    calendarId: calendarId,
    syncToken: storedSyncToken
  };

  var events = Calendar.Events.list(syncParams);

  // Procesar los eventos actualizados
  if (events.updated.length > 0) {
    // Realizar acciones según los eventos actualizados
    for (var i = 0; i < events.updated.length; i++) {
      var updatedEvent = events.updated[i];
      crearNotificacion(updatedEvent);
      Logger.log('Evento actualizado: ' + updatedEvent.summary);
    }

    // Actualizar el token de sincronización almacenado
    PropertiesService.getUserProperties().setProperty(calendarId + '_syncToken', events.nextSyncToken);
  }
}

function zip(...arrays) {
  // Encuentra la longitud mínima de los arreglos para evitar desbordamientos
  const minLength = Math.min(...arrays.map(arr => arr.length));
  return Array.from({ length: minLength }, (_, i) => arrays.map(arr => arr[i]));
}


function prueba() {
  const hoja = excel.getSheetByName('2024')
  Logger.log(hoja.getRange(4,3).getBackground())
}
