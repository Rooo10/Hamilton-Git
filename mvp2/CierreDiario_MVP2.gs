// ============================================================
// CierreDiario_MVP2.gs
// Cierre de caja diario — Hamilton Deco POS MVP2
// Todas las fechas se guardan como fecha real (dd/mm/yyyy),
// misma convención que VENTAS, PAGOS, CC, ITEMS.
// Reutiliza helpers de Dashboard_MVP2.gs: _fechaKey_D, _normMetodo_D,
// COMISION_MAP_D, _leerComisiones_D, METODOS_DASH.
// ============================================================

var SS_ID_CIERRE = '18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8';
var TZ_CIERRE = 'America/Argentina/Buenos_Aires';
var SHEET_CIERRE = 'CIERRE_DIARIO';

var FORMATO_INDICADORES_CD = {
  'Total Ventas del Día': '$#,##0',
  'Cantidad de Ventas': '0',
  'Total $ Descuentos': '$#,##0',
  '% Descuentos': '0.0%',
  'Total a Facturar': '$#,##0',
  '% a Facturar': '0.0%',
  'Cant. Facturadas': '0',
  'Cant. Pend. a Facturar': '0',
  'Total Comisiones del Día': '$#,##0',
  'Comisiones Acumuladas (mes)': '$#,##0',
  'Total Pagos por Ventas': '$#,##0',
  'Total Pagos contra CC': '$#,##0',
  'Saldo CC Sin Vencer': '$#,##0',
  'Saldo CC Vencido': '$#,##0'
};
var LABELS_INDICADORES_CD = Object.keys(FORMATO_INDICADORES_CD);

var FILA_HEADER_FECHAS_CD = 3;
var FILA_PRIMER_INDICADOR_CD = 4;
var FILA_HEADER_ITEMS_CD = FILA_PRIMER_INDICADOR_CD + LABELS_INDICADORES_CD.length + 2;
var FILA_PRIMER_ITEM_CD = FILA_HEADER_ITEMS_CD + 1;

// ── ENTRY POINT: un solo día ──────────────────────────────
function calcularCierreDiario(fechaParam) {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var fecha = fechaParam || new Date();
  var fechaKey = Utilities.formatDate(fecha, TZ_CIERRE, 'yyyy-MM-dd');
  var sh = _asegurarHojaCierre_CD(ss);

  var ventas   = _leerHoja_CD(ss, 'VENTAS', 18);
  var pagos    = _leerHoja_CD(ss, 'PAGOS', 8);
  var pagosCC  = _leerHoja_CD(ss, 'PAGOS_CC', 4);
  var facturas = _leerHoja_CD(ss, 'FACTURAS_PENDIENTES', 15);
  var items    = _leerHoja_CD(ss, 'ITEMS', 8);
  var comMap   = _leerComisiones_D(ss);

  var resultado = _procesarUnDia_CD(sh, fecha, fechaKey, ventas, pagos, pagosCC, facturas, items, comMap, true);
  Logger.log('✅ Cierre diario ' + resultado.fecha + ' | Ventas: $' + resultado.totalVentas +
    ' (' + resultado.cantVentas + ') | Items: ' + resultado.itemsCount);
  return resultado;
}

// ── ENTRY POINT: backfill histórico (solo Bloque A) ────────
function backfillCierreDiario_CD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = _asegurarHojaCierre_CD(ss);

  var ventas   = _leerHoja_CD(ss, 'VENTAS', 18);
  var pagos    = _leerHoja_CD(ss, 'PAGOS', 8);
  var pagosCC  = _leerHoja_CD(ss, 'PAGOS_CC', 4);
  var facturas = _leerHoja_CD(ss, 'FACTURAS_PENDIENTES', 15);
  var items    = _leerHoja_CD(ss, 'ITEMS', 8);
  var comMap   = _leerComisiones_D(ss);

  var fechasSet = {};
  ventas.forEach(function(r) { var k = _fechaKey_D(r[2]); if (k) fechasSet[k] = true; });
  pagos.forEach(function(r) { var k = _fechaKey_D(r[2]); if (k) fechasSet[k] = true; });
  pagosCC.forEach(function(r) { var k = _fechaKey_D(r[2]); if (k) fechasSet[k] = true; });
  facturas.forEach(function(r) { var k = _fechaKey_D(r[1]); if (k) fechasSet[k] = true; });

  var fechasOrdenadas = Object.keys(fechasSet).sort();
  if (fechasOrdenadas.length === 0) { Logger.log('⚠️ No hay fechas con actividad.'); return; }

  Logger.log('Backfill: ' + fechasOrdenadas.length + ' días, desde ' +
    fechasOrdenadas[0] + ' hasta ' + fechasOrdenadas[fechasOrdenadas.length - 1]);

  fechasOrdenadas.forEach(function(fechaKey, idx) {
    var p = fechaKey.split('-');
    var fechaObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
    _procesarUnDia_CD(sh, fechaObj, fechaKey, ventas, pagos, pagosCC, facturas, items, comMap, false);
    if ((idx + 1) % 10 === 0) Logger.log('... ' + (idx + 1) + '/' + fechasOrdenadas.length + ' días');
  });

  Logger.log('✅ Backfill completo: ' + fechasOrdenadas.length + ' días (Bloque A). Saldo CC queda en 0 para fechas pasadas — no hay forma de reconstruir el histórico real.');
}

// ── LÓGICA COMPARTIDA ────────────────────────────────────────
function _procesarUnDia_CD(sh, fecha, fechaKey, ventas, pagos, pagosCC, facturas, items, comMap, escribirItems) {
  var anioNum = parseInt(Utilities.formatDate(fecha, TZ_CIERRE, 'yyyy'));
  var mesNum  = parseInt(Utilities.formatDate(fecha, TZ_CIERRE, 'MM')) - 1;
  var diaNum  = parseInt(Utilities.formatDate(fecha, TZ_CIERRE, 'dd'));
  var fechaSolo = new Date(anioNum, mesNum, diaNum); // fecha real, sin hora

  var totalVentas = 0, cantVentas = 0, totalFacturar = 0;
  ventas.forEach(function(r) {
    if (_fechaKey_D(r[2]) !== fechaKey) return;
    if (String(r[17] || '').toUpperCase() === 'ANULADA') return;
    var total = parseFloat(r[8]) || 0;
    var fact  = String(r[10] || '').toUpperCase();
    cantVentas++;
    totalVentas += total;
    if (fact === 'SI') totalFacturar += total;
  });
  var pctFacturar = totalVentas > 0 ? (totalFacturar / totalVentas) : 0;

  var cantEmitidas = 0, cantPendientes = 0;
  facturas.forEach(function(r) {
    if (_fechaKey_D(r[1]) !== fechaKey) return;
    var estado = String(r[14] || '').toUpperCase();
    if (estado === 'EMITIDA') cantEmitidas++;
    if (estado === 'PENDIENTE') cantPendientes++;
  });

  var totalComisiones = 0, totalPagosVentas = 0;
  pagos.forEach(function(r) {
    if (_fechaKey_D(r[2]) !== fechaKey) return;
    if (String(r[7] || '').toUpperCase() === 'ANULADA') return;
    var monto = parseFloat(r[4]) || 0;
    var metodo = _normMetodo_D(String(r[5] || ''));
    totalPagosVentas += monto;
    if (metodo) {
      var claveParam = COMISION_MAP_D[metodo];
      var tasa = claveParam ? (comMap[claveParam] || 0) : 0;
      totalComisiones += monto * tasa;
    }
  });

  var totalPagosCC = 0;
  pagosCC.forEach(function(r) {
    if (_fechaKey_D(r[2]) !== fechaKey) return;
    totalPagosCC += parseFloat(r[3]) || 0;
  });

  var itemsHoy = [];
  var totalDescuentos = 0;
  items.forEach(function(r) {
    if (_fechaKey_D(r[1]) !== fechaKey) return;
    if (String(r[7] || '').toUpperCase() === 'ANULADA') return;
    var descripcion = String(r[3] || '');
    var subtotal = parseFloat(r[6]) || 0;
    if (descripcion.toUpperCase() === 'DESCUENTO') totalDescuentos += subtotal;
    itemsHoy.push([fechaSolo, r[2], r[3], r[4], subtotal]);
  });
  totalDescuentos = Math.abs(totalDescuentos);
  var pctDescuentos = totalVentas > 0 ? (totalDescuentos / totalVentas) : 0;

  // ── Saldo CC actual (solo tiene sentido para el cierre de HOY;
  // en un backfill de fechas pasadas, LISTADO_CC igual refleja el
  // estado ACTUAL, no el histórico — por eso el backfill llama a
  // esta misma función pero el valor será el mismo en todas las
  // columnas pasadas. Se deja así a propósito: no hay forma de
  // reconstruir el saldo CC histórico real). ──
  var saldoSinVencer = 0, saldoVencido = 0;
  var shListadoCC = sh.getParent().getSheetByName('LISTADO_CC');
  if (shListadoCC && shListadoCC.getLastRow() >= 3) {
    var datosListadoCC = shListadoCC.getRange(3, 5, shListadoCC.getLastRow() - 2, 2).getValues();
    datosListadoCC.forEach(function(r) {
      saldoSinVencer += parseFloat(r[0]) || 0;
      saldoVencido += parseFloat(r[1]) || 0;
    });
  }

  var col = _obtenerColumnaFecha_CD(sh, fechaSolo, fechaKey);

  var valoresPorLabel = {
    'Total Ventas del Día': totalVentas,
    'Cantidad de Ventas': cantVentas,
    'Total $ Descuentos': totalDescuentos,
    '% Descuentos': pctDescuentos,
    'Total a Facturar': totalFacturar,
    '% a Facturar': pctFacturar,
    'Cant. Facturadas': cantEmitidas,
    'Cant. Pend. a Facturar': cantPendientes,
    'Total Comisiones del Día': totalComisiones,
    'Total Pagos por Ventas': totalPagosVentas,
    'Total Pagos contra CC': totalPagosCC,
    'Saldo CC Sin Vencer': saldoSinVencer,
    'Saldo CC Vencido': saldoVencido
  };

  LABELS_INDICADORES_CD.forEach(function(label) {
    if (valoresPorLabel.hasOwnProperty(label)) {
      var fila = _filaIndicador_CD(label);
      sh.getRange(fila, col).setValue(valoresPorLabel[label]);
      sh.getRange(fila, col).setNumberFormat(FORMATO_INDICADORES_CD[label]);
    }
  });

  var filaComision = _filaIndicador_CD('Total Comisiones del Día');
  var filaComisionAcum = _filaIndicador_CD('Comisiones Acumuladas (mes)');
  var colInicioMes = _obtenerPrimeraColumnaDelMes_CD(sh, fecha, col);
  var colLetraInicio = _colALetra_CD(colInicioMes);
  var colLetraFin = _colALetra_CD(col);
  sh.getRange(filaComisionAcum, col).setFormula(
    '=SUM(' + colLetraInicio + filaComision + ':' + colLetraFin + filaComision + ')'
  );
  sh.getRange(filaComisionAcum, col).setNumberFormat('$#,##0');

  if (escribirItems) _escribirItemsDia_CD(sh, itemsHoy);

  var fechaDisplay = Utilities.formatDate(fecha, TZ_CIERRE, 'dd/MM/yyyy');
  return {
    fecha: fechaDisplay, totalVentas: totalVentas, cantVentas: cantVentas,
    totalDescuentos: totalDescuentos, pctDescuentos: pctDescuentos,
    totalFacturar: totalFacturar, pctFacturar: pctFacturar,
    cantEmitidas: cantEmitidas, cantPendientes: cantPendientes,
    totalComisiones: totalComisiones, totalPagosVentas: totalPagosVentas,
    totalPagosCC: totalPagosCC, saldoSinVencer: saldoSinVencer,
    saldoVencido: saldoVencido, itemsCount: itemsHoy.length
  };
}

function _escribirItemsDia_CD(sh, itemsHoy) {
  var ultimaFila = sh.getLastRow();
  if (ultimaFila >= FILA_PRIMER_ITEM_CD) {
    sh.getRange(FILA_PRIMER_ITEM_CD, 1, ultimaFila - FILA_PRIMER_ITEM_CD + 1, 5).clearContent();
  }
  if (itemsHoy.length > 0) {
    sh.getRange(FILA_PRIMER_ITEM_CD, 1, itemsHoy.length, 5).setValues(itemsHoy);
    sh.getRange(FILA_PRIMER_ITEM_CD, 1, itemsHoy.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(FILA_PRIMER_ITEM_CD, 5, itemsHoy.length, 1).setNumberFormat('$#,##0');
  }
}

function limpiarItemsAcumulados_CD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = ss.getSheetByName(SHEET_CIERRE);
  if (!sh) { Logger.log('No existe la hoja CIERRE_DIARIO'); return; }
  var ultimaFila = sh.getLastRow();
  if (ultimaFila >= FILA_PRIMER_ITEM_CD) {
    sh.getRange(FILA_PRIMER_ITEM_CD, 1, ultimaFila - FILA_PRIMER_ITEM_CD + 1, 5).clearContent();
    Logger.log('✅ Bloque de items limpiado (filas ' + FILA_PRIMER_ITEM_CD + ' a ' + ultimaFila + ')');
  } else {
    Logger.log('Nada que limpiar.');
  }
}

// ── HELPERS ───────────────────────────────────────────────
function _filaIndicador_CD(label) {
  return FILA_PRIMER_INDICADOR_CD + LABELS_INDICADORES_CD.indexOf(label);
}

function _asegurarHojaCierre_CD(ss) {
  var sh = ss.getSheetByName(SHEET_CIERRE);
  if (sh) return sh;

  sh = ss.insertSheet(SHEET_CIERRE);

  sh.getRange(1, 1).setValue('📊 CIERRE DIARIO — HAMILTON DECO')
    .setFontWeight('bold').setFontSize(13)
    .setBackground('#1a1a2e').setFontColor('#f0c040');

  sh.getRange(FILA_HEADER_FECHAS_CD, 1).setValue('INDICADOR')
    .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  LABELS_INDICADORES_CD.forEach(function(label, i) {
    sh.getRange(FILA_PRIMER_INDICADOR_CD + i, 1).setValue(label).setFontWeight('bold');
  });

  sh.getRange(FILA_HEADER_ITEMS_CD - 1, 1).setValue('📦 DETALLE DE ITEMS — ÚLTIMO CIERRE')
    .setFontWeight('bold').setBackground('#4a148c').setFontColor('#ffffff');

  sh.getRange(FILA_HEADER_ITEMS_CD, 1, 1, 5)
    .setValues([['FECHA', 'NOMBRE_CLIENTE', 'DESCRIPCION', 'CANTIDAD', 'SUBTOTAL']])
    .setFontWeight('bold').setBackground('#e1bee7');

  sh.setColumnWidth(1, 220);
  sh.setFrozenColumns(1);
  sh.setFrozenRows(FILA_HEADER_FECHAS_CD);

  return sh;
}

function _leerHoja_CD(ss, nombreHoja, numCols) {
  var sh = ss.getSheetByName(nombreHoja);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, numCols).getValues()
    .filter(function(r) { return r[0]; });
}

// Busca la columna cuya fecha real coincida con fechaKey (yyyy-MM-dd).
// Si no existe, crea una columna nueva con una FECHA REAL (no texto).
function _obtenerColumnaFecha_CD(sh, fechaSolo, fechaKey) {
  var headers = sh.getRange(FILA_HEADER_FECHAS_CD, 1, 1, sh.getMaxColumns()).getValues()[0];
  var ultimaColUsada = 1;
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] !== '' && headers[c] !== null) {
      var key = _fechaKey_D(headers[c]);
      if (key === fechaKey) return c + 1;
      ultimaColUsada = c + 1;
    }
  }
  var nuevaCol = ultimaColUsada + 1;
  if (nuevaCol < 2) nuevaCol = 2;
  sh.getRange(FILA_HEADER_FECHAS_CD, nuevaCol).setValue(fechaSolo)
    .setNumberFormat('dd/mm/yyyy')
    .setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('center');
  sh.setColumnWidth(nuevaCol, 90);
  return nuevaCol;
}

function _obtenerPrimeraColumnaDelMes_CD(sh, fecha, colHoy) {
  var mesActualKey = Utilities.formatDate(fecha, TZ_CIERRE, 'yyyy-MM');
  var headers = sh.getRange(FILA_HEADER_FECHAS_CD, 2, 1, colHoy - 1).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    var key = _fechaKey_D(headers[c]);
    if (key && key.substring(0, 7) === mesActualKey) return c + 2;
  }
  return colHoy;
}

function _colALetra_CD(col) {
  var letra = '';
  while (col > 0) {
    var resto = (col - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    col = Math.floor((col - 1) / 26);
  }
  return letra;
}

// ── MIGRACIÓN — correr UNA sola vez ─────────────────────────
// Convierte los encabezados viejos (texto "dd_MM_yy") de las 63
// columnas ya generadas a fecha real dd/mm/yyyy.
function normalizarFechasCierreDiario_CD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = ss.getSheetByName(SHEET_CIERRE);
  if (!sh) { Logger.log('❌ No existe CIERRE_DIARIO'); return; }

  var ultimaCol = sh.getLastColumn();
  if (ultimaCol < 2) { Logger.log('Nada que normalizar.'); return; }

  var headers = sh.getRange(FILA_HEADER_FECHAS_CD, 2, 1, ultimaCol - 1).getValues()[0];
  var nuevos = headers.map(function(val) {
    if (val instanceof Date) return val;
    var s = String(val || '').trim();
    var m = s.match(/^(\d{1,2})_(\d{1,2})_(\d{2,4})$/); // dd_MM_yy(yy)
    if (!m) return val;
    var dia = parseInt(m[1]), mes = parseInt(m[2]);
    var anio = parseInt(m[3]);
    if (anio < 100) anio += 2000;
    return new Date(anio, mes - 1, dia);
  });

  sh.getRange(FILA_HEADER_FECHAS_CD, 2, 1, nuevos.length).setValues([nuevos]);
  sh.getRange(FILA_HEADER_FECHAS_CD, 2, 1, nuevos.length).setNumberFormat('dd/mm/yyyy');

  Logger.log('✅ Encabezados normalizados a dd/mm/yyyy en CIERRE_DIARIO (' + nuevos.length + ' columnas)');
}

// ── TESTING ───────────────────────────────────────────────
function test_calcularCierreDiario() {
  var resultado = calcularCierreDiario(new Date(2026, 4, 21));
  Logger.log(JSON.stringify(resultado, null, 2));
}

// ============================================================
// VENTAS_DIARIAS: Tabla Mensual + Tabla Diaria por forma de pago
// ============================================================

// ← v2.22: ubica las 2 tablas de VENTAS_DIARIAS por su encabezado (la celda "Mes" de la
// columna B) para NO depender de números de fila fijos. Así se pueden borrar o insertar
// filas arriba de las tablas sin romper nada.
// Devuelve { filaMensual, filaDiaria } = primera fila de DATOS de cada tabla, o null si no las encuentra.
function _ubicarTablas_VD(sh) {
  var ultima = sh.getLastRow();
  if (ultima < 1) return null;
  var colB = sh.getRange(1, 2, ultima, 1).getValues();
  var encabezados = [];
  for (var i = 0; i < colB.length; i++) {
    if (String(colB[i][0]).trim().toLowerCase() === 'mes') encabezados.push(i + 1);
  }
  if (encabezados.length < 2) return null;
  return { filaMensual: encabezados[0] + 1, filaDiaria: encabezados[1] + 1 };
}

function completarVentasDiarias_VD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = ss.getSheetByName('VENTAS_DIARIAS');
  if (!sh) { Logger.log('❌ No existe la hoja VENTAS_DIARIAS'); return; }

  var pos = _ubicarTablas_VD(sh);
  if (!pos) { Logger.log('❌ No encontré los 2 encabezados "Mes" (columna B) de las tablas en VENTAS_DIARIAS'); return; }
  var nMetodos = METODOS_DASH.length;

  var pagos = _leerHoja_CD(ss, 'PAGOS', 8);
  var mensual = {};
  for (var m = 0; m < 12; m++) mensual[m] = {};
  var diario = {};

  pagos.forEach(function(r) {
    if (String(r[7] || '').toUpperCase() === 'ANULADA') return;
    var key = _fechaKey_D(r[2]);
    if (!key) return;
    var mesIdx = parseInt(key.substring(5, 7)) - 1;
    var metodo = _normMetodo_D(String(r[5] || ''));
    if (!metodo) return;
    var monto = parseFloat(r[4]) || 0;

    mensual[mesIdx][metodo] = (mensual[mesIdx][metodo] || 0) + monto;
    if (!diario[key]) diario[key] = {};
    diario[key][metodo] = (diario[key][metodo] || 0) + monto;
  });

  // Tabla Mensual: 12 filas debajo de su encabezado
  for (var i = 0; i < 12; i++) {
    var fila = pos.filaMensual + i;
    var valoresMes = METODOS_DASH.map(function(mm) { return mensual[i][mm.key] || 0; });
    sh.getRange(fila, 3, 1, nMetodos).setValues([valoresMes]);
    sh.getRange(fila, 3, 1, nMetodos).setNumberFormat('$#,##0');
  }

  // Tabla Diaria: desde debajo de su encabezado hasta la última fila
  var filaInicio = pos.filaDiaria;
  var ultimaFila = sh.getLastRow();
  var numFilas = ultimaFila - filaInicio + 1;
  if (numFilas > 0) {
    var fechasCol = sh.getRange(filaInicio, 1, numFilas, 1).getValues();
    var filasAEscribir = fechasCol.map(function(row) {
      var key = _fechaKey_D(row[0]);
      var datosDia = diario[key] || {};
      return METODOS_DASH.map(function(mm) { return datosDia[mm.key] || 0; });
    });
    sh.getRange(filaInicio, 3, filasAEscribir.length, nMetodos).setValues(filasAEscribir);
    sh.getRange(filaInicio, 3, filasAEscribir.length, nMetodos).setNumberFormat('$#,##0');
  }

  Logger.log('✅ VENTAS_DIARIAS completada: Tabla Mensual (filas ' + pos.filaMensual + '-' + (pos.filaMensual + 11) +
             ') + Tabla Diaria (' + numFilas + ' filas desde la ' + filaInicio + ')');
}

// ── MIGRACIÓN — correr UNA sola vez ─────────────────────────
// Convierte la Tabla Diaria de VENTAS_DIARIAS (col A, dd/mm/yy con
// año de 2 dígitos) a fecha real dd/mm/yyyy.
function normalizarFechasVentasDiarias_VD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = ss.getSheetByName('VENTAS_DIARIAS');
  if (!sh) { Logger.log('❌ No existe VENTAS_DIARIAS'); return; }

  var filaInicio = 33;
  var ultimaFila = sh.getLastRow();
  var numFilas = ultimaFila - filaInicio + 1;
  if (numFilas <= 0) { Logger.log('Nada que normalizar.'); return; }

  var celdas = sh.getRange(filaInicio, 1, numFilas, 1).getValues();
  var nuevas = celdas.map(function(row) {
    var val = row[0];
    if (val instanceof Date) return [val];
    var s = String(val || '').trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (!m) return [val];
    var dia = parseInt(m[1]), mes = parseInt(m[2]);
    var anio = parseInt(m[3]);
    if (anio < 100) anio += 2000;
    return [new Date(anio, mes - 1, dia)];
  });

  sh.getRange(filaInicio, 1, nuevas.length, 1).setValues(nuevas);
  sh.getRange(filaInicio, 1, nuevas.length, 1).setNumberFormat('dd/mm/yyyy');

  Logger.log('✅ Fechas normalizadas a dd/mm/yyyy en ' + nuevas.length + ' filas (VENTAS_DIARIAS)');
}

// Agrega la fórmula =SUM(C:Q) en la columna Total (R) de ambas
// tablas de VENTAS_DIARIAS. Es una fórmula, no un valor: se
// recalcula sola si los datos de esa fila cambian.
function agregarTotalVentasDiarias_VD() {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = ss.getSheetByName('VENTAS_DIARIAS');
  if (!sh) { Logger.log('❌ No existe VENTAS_DIARIAS'); return; }

  // Tabla Mensual: filas 18-29
  for (var fila = 18; fila <= 29; fila++) {
    sh.getRange(fila, 18).setFormula('=SUM(C' + fila + ':Q' + fila + ')');
    sh.getRange(fila, 18).setNumberFormat('$#,##0');
  }

  // Tabla Diaria: desde fila 33 hasta la última fila con datos
  var filaInicio = 33;
  var ultimaFila = sh.getLastRow();
  for (var f = filaInicio; f <= ultimaFila; f++) {
    sh.getRange(f, 18).setFormula('=SUM(C' + f + ':Q' + f + ')');
    sh.getRange(f, 18).setNumberFormat('$#,##0');
  }
  Logger.log('✅ Columna Total agregada: filas 18-29 (mensual) y 33-' + ultimaFila + ' (diaria)');
}
