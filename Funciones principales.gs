const calendario = CalendarApp.getCalendarById('057c4a0768093a7c3ce599e91e3dd992dbf940d5e03ad8878eb8b6742a836062@group.calendar.google.com');
const excel = SpreadsheetApp.openById('1vUOZWTvjnCc2DVkfLgxeFkb6EGGcmk9tqZWdZqGl8ps');
const hoja_resumen = excel.getSheetByName('Resumen');
const colores_meses = ['#e6b8af', '#fce5cd']// Colores alternos para los meses (cambiar manualmente)

function cargarClasesDia(dia) { //Función para cargar las clases de un dado día
  var fecha;
  if (dia instanceof Date && !isNaN(dia.getTime())) {
    // Se proporcionó un argumento válido, usarlo como fecha
    fecha = new Date(dia);
  } else {
    // No se proporcionó un argumento válido o no se proporcionó ningún argumento, usar la fecha actual
    fecha = new Date();
  }
  const hoja = excel.getSheetByName(fecha.getFullYear());
  const fila_fecha = hoja.getLastRow()+1;
  const col_fecha = encontrarColumna(hoja, 'Fecha');
  const col_hora = encontrarColumna(hoja, 'Hora');
  const col_alumne = encontrarColumna(hoja, 'Alumne');
  const col_duracion = encontrarColumna(hoja, 'Duración');
  const col_pesos = encontrarColumna(hoja, 'Pesos');
  const col_dolares = encontrarColumna(hoja, 'Dólares');
  const col_estado = encontrarColumna(hoja, 'Estado del pago');
  const col_precio = encontrarColumna(hoja, 'Precio actual');
  const pesos = hoja.getRange(3,col_precio).getValue();
  const cambio = dolarBlue(fecha)
  const eventos = calendario.getEventsForDay(fecha);
  const colores_fecha = ['#8e7cc3', '#ffd966']// Colores alternos para los días (cambiar manualmente)
  const colores = getColores(alternarColor(conseguirColor(hoja.getRange(fila_fecha-1,col_fecha)), colores_fecha), colores_fecha)

  if (eventos.length != 0) {
    for (var evento of eventos) {
      if (evento.getColor()==10) {continue}
      var alumne = evento.getTitle();
      var hora = evento.getStartTime();
      var duracion = evento.getEndTime() - evento.getStartTime();
      var fila = hoja.getLastRow()+1;
      var columnas = [col_hora, col_alumne, col_duracion, col_pesos, col_dolares];
      var valores = [hora,alumne,`=${duracion/(24*60*60*1000)}`.replace('.',','),`=${pesos*duracion/(60*60*1000)}`.replace('.',','),`=${pesos*duracion/(60*60*1000)/cambio}`.replace('.',',')];

      zip(columnas,valores).forEach(([columna, valor]) => {
        var celda = hoja.getRange(fila, columna);
        celda.setValue(valor);
        colorAlterno(celda, colores, hoja);
      });
      hoja.getRange(fila, col_estado).insertCheckboxes();
    };

    var celda_fecha = hoja.getRange(fila_fecha,col_fecha,eventos.length,1);
    celda_fecha.setValue(fecha);
    colorAlterno(celda_fecha, colores_fecha, hoja)
    celda_fecha.merge();

    var celda_mes = hoja.getRange(fila_fecha-1,col_fecha-1)
    if (celda_mes.getMergedRanges().length == 0){
      try {
        var mes = celda_mes.getValue().getMonth()
      } catch {
        if (primerEventoAno(fecha)) {
          var mes = NaN
        } else {
        throw 'No hay mes anterior y no es el primer evento del año.'
        }
      };
      var fila_mes = celda_mes.getRow();
    } else {
      var fila_mes = celda_mes.getMergedRanges()[0].getRow();
      var mes = hoja.getRange(fila_mes, col_fecha-1).getValue().getMonth();
    }

    if (fecha.getMonth() == mes) {
      hoja.getRange(fila_mes,col_fecha-1,fila-fila_mes+1,1).merge();

      hoja.getRange(fila_mes,col_duracion+1,fila-fila_mes+1,1).merge();
      var col_letra_duracion = columnToLetter(col_duracion);
      var rango = col_letra_duracion+fila_mes+':'+col_letra_duracion+fila;
      hoja.getRange(fila_mes,col_duracion+1,fila-fila_mes+1,1).setValue(`=SUM(${rango})`);

      hoja.getRange(fila_mes,col_pesos+1,fila-fila_mes+1,1).merge();
      var col_letra_pesos = columnToLetter(col_pesos);
      var rango = col_letra_pesos+fila_mes+':'+col_letra_pesos+fila;
      hoja.getRange(fila_mes,col_pesos+1,fila-fila_mes+1,1).setValue(`=SUM(${rango})`);

      hoja.getRange(fila_mes,col_dolares+1,fila-fila_mes+1,1).merge();
      var col_letra_dolares = columnToLetter(col_dolares);
      var rango = col_letra_dolares+fila_mes+':'+col_letra_dolares+fila;
      hoja.getRange(fila_mes,col_dolares+1,fila-fila_mes+1,1).setValue(`=SUM(${rango})`);

      for (var col of [col_alumne, col_dolares, col_duracion, col_estado, col_fecha, col_hora, col_pesos]) {
        var celda = hoja.getRange(fila_fecha,col);
        if (celda.getMergedRanges().length == 0) {
          celda.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
        } else {
          celda.getMergedRanges()[0].setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
        };
      };
    } else {
      //mes
      var celda_estemes = hoja.getRange(fila_fecha, col_fecha-1, eventos.length, 1);
      celda_estemes.setValue(fecha);
      colorAlterno(celda_estemes, colores_meses, hoja)
      celda_estemes.merge();
      var fila_final = fila_fecha + eventos.length - 1

      //duracion total
      var celda_duracion = hoja.getRange(fila_fecha, col_duracion+1, eventos.length, 1);
      var col_letra_duraciones = columnToLetter(col_duracion);
      var rango = col_letra_duraciones+fila_fecha+':'+col_letra_duraciones+fila_final;
      celda_duracion.setValue(`=SUM(${rango})`);
      colorAlterno(celda_duracion, colores_meses, hoja)
      celda_duracion.merge();

      //pesos totales
      var celda_pesos = hoja.getRange(fila_fecha, col_pesos+1, eventos.length, 1);
      var col_letra_pesoss = columnToLetter(col_pesos);
      var rango = col_letra_pesoss+fila_fecha+':'+col_letra_pesoss+fila_final;
      celda_pesos.setValue(`=SUM(${rango})`);
      colorAlterno(celda_pesos, colores_meses, hoja)
      celda_pesos.merge();

      //dólares totales
      var celda_dolares = hoja.getRange(fila_fecha, col_dolares+1, eventos.length, 1);
      var col_letra_dolaress = columnToLetter(col_dolares);
      var rango = col_letra_dolaress+fila_fecha+':'+col_letra_dolaress+fila_final;
      celda_dolares.setValue(`=SUM(${rango})`);
      colorAlterno(celda_dolares, colores_meses, hoja)
      celda_dolares.merge();
    }

    for (var col of [col_alumne, col_dolares, col_dolares+1, col_duracion, col_duracion+1, col_estado, col_hora, col_pesos, col_pesos+1, col_fecha, col_fecha-1]) {
      var celda = hoja.getRange(fila+1,col);
      celda.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    }
  }
}

function cargarMes() { //Función para cargar el resumen del mes
  const hoy = new Date();
  const mes = hoy.getMonth()-1;
  let ano;
  if (mes!=-1) {ano = hoy.getFullYear()} else {ano = hoy.getFullYear()-1}
  const cambio = dolarBlueMes(ano, mes);
  const hoja_ano = excel.getSheetByName(new String(ano));
  const celda = hoja_ano.getRange(hoja_ano.getLastRow(), encontrarColumna(hoja_ano, 'Fecha')-1);
  let fila_mes
  if (celda.getMergedRanges().length!=0) {
    fila_mes = celda.getMergedRanges()[0].getRow();
  } else {
    fila_mes = celda.getRow();
  }
  
  const fila = hoja_resumen.getLastRow()+1;
  hoja_resumen.getRange(fila,encontrarColumna(hoja_resumen, 'Mes')).setValue(new Date(ano, mes));
  const total_pesos = hoja_ano.getRange(fila_mes,encontrarColumna(hoja_ano,'Pesos')+1).getValue();
  hoja_resumen.getRange(fila,encontrarColumna(hoja_resumen, 'Pesos')).setValue(total_pesos);
  const total_dolares = hoja_ano.getRange(fila_mes,encontrarColumna(hoja_ano,'Dólares')+1).getValue();
  hoja_resumen.getRange(fila,encontrarColumna(hoja_resumen, 'Dólares')).setValue(total_dolares);
  const total_horas = getDuracion(hoja_ano.getRange(fila_mes,encontrarColumna(hoja_ano,'Duración')+1));
  hoja_resumen.getRange(fila,encontrarColumna(hoja_resumen, 'Horas')).setValue(total_horas);
}

function cargarCanasta() {
  const col_canasta = encontrarColumna(hoja_resumen, 'Canasta básica alimentaria');
  let fila=4;
  while (true) {
    if (hoja_resumen.getRange(fila,col_canasta).getValue()==String()) {
      break
    } else {
      fila++
    }
  }
  mes = hoja_resumen.getRange(fila,encontrarColumna(hoja_resumen,'Mes')).getValue();
  let canasta;
  try {
    canasta = canastaBasica(mes);
  } catch (error) {
    Logger.log(error);
    return
  }
  const total_pesos = hoja_resumen.getRange(fila, encontrarColumna(hoja_resumen, 'Pesos')).getValue();
  hoja_resumen.getRange(fila, encontrarColumna(hoja_resumen, 'Pesos')).setNumberFormat('0.00');
  const hora_pesos = hoja_resumen.getRange(fila, encontrarColumna(hoja_resumen, 'Pesos')+1).getValue();
  hoja_resumen.getRange(fila, encontrarColumna(hoja_resumen, 'Pesos')+1).setNumberFormat('0.00');

  hoja_resumen.getRange(fila,col_canasta).setValue(total_pesos/canasta);
  hoja_resumen.getRange(fila, col_canasta+1).setValue(hora_pesos/canasta);
}

function crearNotificaciones() {
  const dia_eventos = new Date();
  if (dia_eventos.getHours()>14) {dia_eventos.setDate(dia_eventos.getDate() + 1)};
  const tiempo_notificacion = new Date();
  tiempo_notificacion.setHours(tiempo_notificacion.getHours+1)
  tiempo_notificacion.setMinutes(0)
  tiempo_notificacion.setSeconds(0)
  tiempo_notificacion.setMilliseconds(0)  
  const eventos  = calendario.getEventsForDay(dia_eventos);
  for (var evento of eventos) {
    const fecha = evento.getStartTime();
    const intervalo_minutos = (fecha - tiempo_notificacion)/1000/60;
    evento.addPopupReminder(intervalo_minutos);
  }
}

function PRUEBA() { //Función de prueba
  const fechaInicio = new Date(2024, 6, 31);  // 1 de enero del año especificado
  const fechaFin = new Date(2024, 7, 1); // 1 de enero del año siguiente

  for (let fecha = fechaInicio; fecha < fechaFin; fecha.setDate(fecha.getDate() + 1)) {
    cargarClasesDia(fecha);
    Logger.log(fecha);
  }
}
