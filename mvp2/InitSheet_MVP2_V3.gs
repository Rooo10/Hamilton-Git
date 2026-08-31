// ============================================================
// InitSheet_MVP2_V3.gs  —  ejecutar inicializarSheet_V3() UNA VEZ
// Agrega sección COMISIONES y PLAZOS en hoja PARAMETROS
// ============================================================

var SPREADSHEET_ID_INIT = '18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8';

function inicializarSheet_V3() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_INIT);
  _actualizarParametros_V3(ss);
  Logger.log('✅ PARAMETROS actualizado con COMISIONES y PLAZOS.');
}

function _actualizarParametros_V3(ss) {
  var sheet = ss.getSheetByName('PARAMETROS');
  if (!sheet) { Logger.log('❌ No se encontró PARAMETROS.'); return; }

  var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');

  // ── Fila 9: separador + header histórico ──────────────────
  sheet.getRange('A9').setValue('— COMISIONES —')
       .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  sheet.getRange('B9').setValue('VALOR_ACTUAL')
       .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#f0c040');
  sheet.getRange('C9').setValue(hoy)
       .setFontWeight('bold').setBackground('#c8e6c9').setFontColor('#1b5e20')
       .setNote('Primera carga. Para actualizar: agregar col D con nueva fecha y nuevos valores. Pintar en verde. El sistema siempre lee col B.');

  // ── Filas 10-23: comisiones ───────────────────────────────
  var comisiones = [
    ['COMISION_EFECTIVO',           0,      'Sin costo'],
    ['COMISION_TRANSFERENCIA',      0,      'Transferencia bancaria CBU/alias — sin costo'],
    ['COMISION_MERCADO_PAGO',       0.0299, 'Mercado Pago QR'],
    ['COMISION_QR_PCT',             0.008,  'QR Payway (PCT)'],
    ['COMISION_LINK_PAGO',          0.030,  'Link de Pago / Venta Online'],
    ['COMISION_VISA_DEBITO',        0.012,  'Visa Débito — terminal Payway'],
    ['COMISION_MASTERCARD_DEBITO',  0.014,  'Mastercard Débito — terminal Payway'],
    ['COMISION_CABAL_DEBITO',       0.012,  'Cabal Débito'],
    ['COMISION_VISA_CREDITO',       0.020,  'Visa Crédito — 1 pago y cuotas'],
    ['COMISION_MASTERCARD_CREDITO', 0.020,  'Mastercard Crédito'],
    ['COMISION_CABAL_CREDITO',      0.020,  'Cabal Crédito'],
    ['COMISION_AMEX',               0.028,  'American Express'],
    ['COMISION_DOLARES',            0,      'Dólares billete — sin costo procesador'],
    ['COMISION_CUENTA_CORRIENTE',   0,      'Deuda interna — sin costo procesador'],
  ];

  for (var i = 0; i < comisiones.length; i++) {
    var fila = 10 + i;
    var c = comisiones[i];
    sheet.getRange(fila, 1).setValue(c[0]);
    sheet.getRange(fila, 2).setValue(c[1]).setNumberFormat('0.00%').setBackground('#fff9c4');
    sheet.getRange(fila, 3).setValue(c[1]).setNumberFormat('0.00%')
         .setBackground('#c8e6c9').setNote(c[2]);
  }

  // ── Fila 25: separador PLAZOS ─────────────────────────────
  sheet.getRange('A25').setValue('— PLAZOS ACREDITACION (días hábiles) —')
       .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  sheet.getRange('B25').setValue('VALOR_ACTUAL')
       .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#f0c040');
  sheet.getRange('C25').setValue(hoy)
       .setFontWeight('bold').setBackground('#c8e6c9').setFontColor('#1b5e20');

  // ── Filas 26-34: plazos ───────────────────────────────────
  var plazos = [
    ['PLAZO_DEBITO',               1,  'Visa/Master/Cabal débito — 1 día hábil'],
    ['PLAZO_CREDITO_1PAGO',        8,  'Crédito 1 pago — comercio pequeño BCRA'],
    ['PLAZO_CREDITO_CUOTAS_FIJAS', 2,  'Cuotas Fijas Payway'],
    ['PLAZO_CREDITO_CUOTAS_PW',    1,  'Cuotas Payway — 1 día hábil'],
    ['PLAZO_CREDITO_MIPYME',       10, 'Cuotas MiPyME — 10 días hábiles'],
    ['PLAZO_AMEX_1PAGO',           9,  'Amex 1 pago'],
    ['PLAZO_QR_PCT',               0,  'QR Payway PCT — inmediato'],
    ['PLAZO_MERCADO_PAGO',         1,  'Mercado Pago — 1 día hábil'],
    ['PLAZO_LINK_PAGO',            2,  'Link de Pago — estimado'],
  ];

  for (var j = 0; j < plazos.length; j++) {
    var filaP = 26 + j;
    var p = plazos[j];
    sheet.getRange(filaP, 1).setValue(p[0]);
    sheet.getRange(filaP, 2).setValue(p[1]).setBackground('#fff9c4');
    sheet.getRange(filaP, 3).setValue(p[1]).setBackground('#c8e6c9').setNote(p[2]);
  }

  // ── Instrucciones fila 37 ─────────────────────────────────
  sheet.getRange('A37').setValue('⚠️ CÓMO ACTUALIZAR COMISIONES:')
       .setFontWeight('bold').setFontColor('#b71c1c');
  sheet.getRange('A38').setValue(
    '1) Agregar nueva columna a la derecha de la última. ' +
    '2) Escribir fecha en fila 9 (COMISIONES) y fila 25 (PLAZOS). ' +
    '3) Pegar valores anteriores. 4) Editar col B con valores nuevos. ' +
    '5) Pintar nueva columna en verde. El sistema siempre lee col B.'
  ).setWrap(true).setBackground('#fff3e0');

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  Logger.log('✅ PARAMETROS actualizado');
}
