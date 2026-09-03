var SS_ID_DASH = '18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8';
var TZ_DASH    = 'America/Argentina/Buenos_Aires';
var MESES      = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

var METODOS_DASH = [
  { label: 'Efectivo',         key: 'EFECTIVO' },
  { label: 'Transferencia',    key: 'TRANSFERENCIA' },
  { label: 'Mercado Pago',     key: 'MERCADO_PAGO' },
  { label: 'QR Payway',        key: 'QR_PCT' },
  { label: 'Link de Pago',     key: 'LINK_PAGO' },
  { label: 'Visa Débito',      key: 'VISA_DEBITO' },
  { label: 'Master Débito',    key: 'MASTERCARD_DEBITO' },
  { label: 'Cabal Débito',     key: 'CABAL_DEBITO' },
  { label: 'Visa Crédito',     key: 'VISA_CREDITO' },
  { label: 'Master Crédito',   key: 'MASTERCARD_CREDITO' },
  { label: 'Cabal Crédito',    key: 'CABAL_CREDITO' },
  { label: 'Amex',             key: 'AMEX' },
  { label: 'Tarjeta Naranja',  key: 'TARJETA_NARANJA' },
  { label: 'Dólares',          key: 'DOLARES' },
  { label: 'Cuenta Corriente', key: 'CUENTA_CORRIENTE' },
];

var COMISION_MAP_D = {
  'EFECTIVO':'COMISION_EFECTIVO',
  'TRANSFERENCIA':'COMISION_TRANSFERENCIA',
  'MERCADO_PAGO':'COMISION_MERCADO_PAGO',
  'QR_PCT':'COMISION_QR_PCT',
  'LINK_PAGO':'COMISION_LINK_PAGO',
  'VISA_DEBITO':'COMISION_VISA_DEBITO',
  'MASTERCARD_DEBITO':'COMISION_MASTERCARD_DEBITO',
  'CABAL_DEBITO':'COMISION_CABAL_DEBITO',
  'VISA_CREDITO':'COMISION_VISA_CREDITO',
  'MASTERCARD_CREDITO':'COMISION_MASTERCARD_CREDITO',
  'CABAL_CREDITO':'COMISION_CABAL_CREDITO',
  'AMEX':'COMISION_AMEX',
  'TARJETA_NARANJA':'COMISION_TARJETA_NARANJA',
  'DOLARES':'COMISION_DOLARES',
  'CUENTA_CORRIENTE':'COMISION_CUENTA_CORRIENTE',
};

var ALIAS_METODO_D = {
  'TRANSFERENCIA_QR':'QR_PCT',
  'QR':'QR_PCT',
  'QR_PAYWAY':'QR_PCT',
  'MERCADOPAGO':'MERCADO_PAGO',
  'LINK_DE_PAGO':'LINK_PAGO',
  'CC':'CUENTA_CORRIENTE',
  'DOLARES':'DOLARES',
};

// ── ENTRY POINT ──────────────────────────────────────────────
function generarDashboard() {
  var ss     = SpreadsheetApp.openById(SS_ID_DASH);
  var ahora  = new Date();
  var anio   = parseInt(Utilities.formatDate(ahora, TZ_DASH, 'yyyy'));
  var mes    = parseInt(Utilities.formatDate(ahora, TZ_DASH, 'MM')) - 1;
  var hoyStr = Utilities.formatDate(ahora, TZ_DASH, 'dd/MM/yyyy HH:mm');
  var hoyKey = Utilities.formatDate(ahora, TZ_DASH, 'yyyy-MM-dd');

  var ventas = _leerVentas_D(ss);
  var pagos  = _leerPagos_D(ss);
  var ccData = _leerCC_D(ss);
  var comMap = _leerComisiones_D(ss);
  var agg    = _agregar_D(ventas, pagos, ccData, comMap, anio, mes, hoyKey, ahora);
  _escribirDashboard_D(ss, agg, hoyStr, anio, mes);

  // ← NUEVO: agrega columnas Q/R/S (Total $, Total %, Partida) sobre DASHBOARD
  _agregarColumnasQRS_D(ss, anio);

  Logger.log('✅ Dashboard generado — ' + hoyStr);
}

// ← FIX v2.14: wrapper para el trigger time-based con try/catch
// Apuntar el trigger a esta función, NO a generarDashboard directamente
function triggerDashboard() {
  try {
    generarDashboard();
  } catch(e) {
    Logger.log('❌ Error en generarDashboard: ' + e.message + ' | Stack: ' + e.stack);
    // No relanza la excepción → GAS no envía email de fallo
  }
}

// ── MENU ─────────────────────────────────────────────────────
function _crearMenuDashboard() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Dashboard')
    .addItem('🔄 Actualizar Dashboard', 'generarDashboard')
    .addToUi();
}

// ── LECTURA ───────────────────────────────────────────────────
function _leerVentas_D(ss) {
  var sh = ss.getSheetByName('VENTAS');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues()
    .filter(function(r){ return r[0] && r[17] !== 'ANULADA'; });
}

function _leerPagos_D(ss) {
  var sh = ss.getSheetByName('PAGOS');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 8).getValues()
    .filter(function(r){ return r[0] && r[7] !== 'ANULADA'; });
}

function _leerCC_D(ss) {
  var sh = ss.getSheetByName('CC');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 12).getValues()
    .filter(function(r){ return r[0]; });
}

function _leerComisiones_D(ss) {
  var sh = ss.getSheetByName('PARAMETROS');
  if (!sh) return {};
  var data = sh.getRange('A10:B24').getValues();
  var map = {};
  data.forEach(function(r){ if(r[0]) map[r[0]] = parseFloat(r[1]) || 0; });
  return map;
}

// ── AGREGACIÓN ────────────────────────────────────────────────
function _agregar_D(ventas, pagos, ccData, comMap, anio, mesActual, hoyKey, ahora) {
  var a = {
    v_hoy_cant:0, v_hoy_monto:0,
    v_fact_hoy_cant:0, v_fact_hoy_monto:0,
    v_mes_cant:_z12_D(), v_mes_monto:_z12_D(),
    v_fact_mes_cant:_z12_D(), v_fact_mes_monto:_z12_D(),
    p:{},
    com_hoy:0, com_mes:_z12_D(),
    net_hoy:0, net_mes:_z12_D(),
    cc_pend_cant:0, cc_pend_monto:0,
    cc_venc_cant:0, cc_venc_monto:0,
    cc_venc7_cant:0, cc_venc7_monto:0,
    cc_nuevas_mes_cant:_z12_D(), cc_nuevas_mes_monto:_z12_D(),
    top:{}, mesActual:mesActual,
  };
  METODOS_DASH.forEach(function(m){ a.p[m.key]={hoy:0,mes:_z12_D()}; });

  ventas.forEach(function(r) {
    var key = _fechaKey_D(r[2]);
    if (!key) return;
    var anioF = parseInt(key.substring(0,4));
    var mesF  = parseInt(key.substring(5,7)) - 1;
    if (anioF !== anio) return;
    var tot     = parseFloat(r[8]) || 0;
    var esFact  = (r[10] === true || String(r[10]).toUpperCase() === 'SI');
    var impFact = parseFloat(r[14]) || 0;
    var nombre  = String(r[5]||'').trim().toUpperCase();
    var esHoy   = (key === hoyKey);

    if (esHoy) {
      a.v_hoy_cant++; a.v_hoy_monto += tot;
      if (esFact) { a.v_fact_hoy_cant++; a.v_fact_hoy_monto += impFact; }
    }
    a.v_mes_cant[mesF]++; a.v_mes_monto[mesF] += tot;
    if (esFact) { a.v_fact_mes_cant[mesF]++; a.v_fact_mes_monto[mesF] += impFact; }
    if (mesF === mesActual && nombre &&
        nombre !== 'CLIENTE MOSTRADOR' && nombre !== 'CONSUMIDOR FINAL') {
      a.top[nombre] = (a.top[nombre]||0) + tot;
    }
  });

  pagos.forEach(function(r) {
    var key = _fechaKey_D(r[2]);
    if (!key) return;
    var anioF = parseInt(key.substring(0,4));
    var mesF  = parseInt(key.substring(5,7)) - 1;
    if (anioF !== anio) return;
    var metodo = _normMetodo_D(String(r[5]||''));
    var monto  = parseFloat(r[4]) || 0;
    var esHoy  = (key === hoyKey);
    if (!metodo) return;

    var clave    = COMISION_MAP_D[metodo] || null;
    var tasa     = clave ? (comMap[clave]||0) : 0;
    var comision = monto * tasa;

    if (esHoy) {
      a.p[metodo].hoy += monto;
      a.com_hoy += comision;
      a.net_hoy += (monto - comision);
    }
    a.p[metodo].mes[mesF] += monto;
    a.com_mes[mesF] += comision;
    a.net_mes[mesF] += (monto - comision);
  });

  ccData.forEach(function(r) {
    var estado = String(r[11]||'').trim().toUpperCase();
    if (estado === 'PAGADO' || estado === 'ANULADA') return;
    var imp    = parseFloat(r[6]) || 0;
    var pagado = parseFloat(r[7]) || 0;
    var saldo  = imp - pagado;
    var keyVenc  = _fechaKey_D(r[8]);
    var keyVenta = _fechaKey_D(r[2]);

    a.cc_pend_cant++; a.cc_pend_monto += saldo;
    if (estado === 'VENCIDO') { a.cc_venc_cant++; a.cc_venc_monto += saldo; }

    if (keyVenc && estado === 'SIN_VENCER') {
      var fVenc = _keyToDate_D(keyVenc);
      var dias  = Math.floor((fVenc - ahora) / 86400000);
      if (dias >= 0 && dias <= 7) { a.cc_venc7_cant++; a.cc_venc7_monto += saldo; }
    }
    if (keyVenta) {
      var anioV = parseInt(keyVenta.substring(0,4));
      var mesV  = parseInt(keyVenta.substring(5,7)) - 1;
      if (anioV === anio) {
        a.cc_nuevas_mes_cant[mesV]++;
        a.cc_nuevas_mes_monto[mesV] += imp;
      }
    }
  });

  return a;
}

// ── ESCRITURA ─────────────────────────────────────────────────
// ← FIX v2.14: reescritura con setValues() batch por fila → elimina timeout
function _escribirDashboard_D(ss, a, hoyStr, anio, mesActual) {
  var sh = ss.getSheetByName('DASHBOARD');
  if (!sh) { Logger.log('❌ Hoja DASHBOARD no encontrada'); return; }
  sh.clearContents(); sh.clearFormats();
  sh.setColumnWidth(1,240); sh.setColumnWidth(2,115);
  for (var ci=0; ci<12; ci++) sh.setColumnWidth(3+ci,88);
  sh.setColumnWidth(15,115);

  var BG_HEAD='#1a1a2e', FG_HEAD='#ffffff', ACCENT='#f0c040';
  var BG_HOY='#e8f5e9',  FG_HOY='#1b5e20';
  var BG_TOT='#fff3e0',  FG_TOT='#e65100';
  var ALT1='#ffffff',    ALT2='#f5f5f5';
  var fila = 1;

  _mf_D(sh,fila,1,1,15,'🏠 HAMILTON DECO · DASHBOARD '+anio,ACCENT,BG_HEAD,13,true); fila++;
  _mf_D(sh,fila,1,1,15,'Actualizado: '+hoyStr,'#888888','#ffffff',9,false); fila++;
  fila++;

  var headerVals = [['']];
  for (var mi=0; mi<12; mi++) headerVals[0].push(MESES[mi]);
  headerVals[0].unshift('');
  var hdrRow = ['', 'HOY'];
  for (var mi=0; mi<12; mi++) hdrRow.push(MESES[mi]);
  hdrRow.push('TOTAL ' + anio);
  sh.getRange(fila, 1, 1, 15).setValues([hdrRow]);
  sh.getRange(fila,2).setFontWeight('bold').setBackground(BG_HOY).setFontColor(FG_HOY).setHorizontalAlignment('center');
  for (var mi=0; mi<12; mi++) {
    sh.getRange(fila,3+mi).setFontWeight('bold')
      .setBackground(mi===mesActual?'#3949ab':BG_HEAD)
      .setFontColor(FG_HEAD).setHorizontalAlignment('center');
  }
  sh.getRange(fila,15).setFontWeight('bold').setBackground(BG_TOT).setFontColor(FG_TOT).setHorizontalAlignment('center');
  fila++;

  // ══ A — VENTAS ══
  fila=_hdr_D(sh,fila,'🛒  VENTAS',BG_HEAD,ACCENT);
  fila=_row_D(sh,fila,'Ventas [cantidad]',a.v_hoy_cant,a.v_mes_cant,'0',ALT1,BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'Ventas [$]',a.v_hoy_monto,a.v_mes_monto,'$#,##0',ALT2,BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'Facturadas ARCA [cantidad]',a.v_fact_hoy_cant,a.v_fact_mes_cant,'0',ALT1,BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'Facturadas ARCA [$]',a.v_fact_hoy_monto,a.v_fact_mes_monto,'$#,##0',ALT2,BG_HOY,BG_TOT,mesActual);
  fila++;

  // ══ B — FORMAS DE PAGO ══
  fila=_hdr_D(sh,fila,'💳  TOTALES POR FORMA DE PAGO',BG_HEAD,ACCENT);
  METODOS_DASH.forEach(function(m,idx){
    var d=a.p[m.key]||{hoy:0,mes:_z12_D()};
    fila=_row_D(sh,fila,m.label,d.hoy,d.mes,'$#,##0',
      (idx%2===0)?ALT1:ALT2,BG_HOY,BG_TOT,mesActual);
  });
  fila++;

  // ══ C — COMISIONES ══
  fila=_hdr_D(sh,fila,'📉  COMISIONES ESTIMADAS','#880e4f','#f8bbd0');
  fila=_row_D(sh,fila,'Total comisiones [$]',a.com_hoy,a.com_mes,'$#,##0','#fce4ec',BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'Total neto estimado [$]',a.net_hoy,a.net_mes,'$#,##0','#fff9c4',BG_HOY,BG_TOT,mesActual);
  sh.getRange(fila,1).setValue('* Calculado sobre tasas en PARAMETROS col B')
    .setFontStyle('italic').setFontColor('#aaaaaa').setFontSize(8);
  fila+=2;

  // ══ D — CUENTAS CORRIENTES ══
  fila=_hdr_D(sh,fila,'📋  CUENTAS CORRIENTES','#0d47a1','#bbdefb');
  fila=_row_D(sh,fila,'CC pendientes [cantidad]',a.cc_pend_cant,a.cc_nuevas_mes_cant,'0',ALT1,BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'CC pendientes [$]',a.cc_pend_monto,a.cc_nuevas_mes_monto,'$#,##0',ALT2,BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'⚠️ CC vencidas HOY [cantidad]',a.cc_venc_cant,_z12_D(),'0','#ffebee',BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'⚠️ CC vencidas HOY [$]',a.cc_venc_monto,_z12_D(),'$#,##0','#ffcdd2',BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'🔔 Vencen próx. 7 días [cant.]',a.cc_venc7_cant,_z12_D(),'0','#fff8e1',BG_HOY,BG_TOT,mesActual);
  fila=_row_D(sh,fila,'🔔 Vencen próx. 7 días [$]',a.cc_venc7_monto,_z12_D(),'$#,##0','#fff3e0',BG_HOY,BG_TOT,mesActual);
  fila++;

  // ══ E — TOP 10 ══
  fila=_hdr_D(sh,fila,'🏆  TOP 10 CLIENTES — '+MESES[mesActual]+' '+anio,'#4a148c','#e1bee7');
  sh.getRange(fila,1,1,2).setValues([['Cliente','Total [$]']]);
  sh.getRange(fila,1).setFontWeight('bold').setBackground('#ce93d8');
  sh.getRange(fila,2).setFontWeight('bold').setBackground('#ce93d8').setHorizontalAlignment('right');
  fila++;
  var top10=Object.keys(a.top).map(function(k){return{n:k,v:a.top[k]};})
    .sort(function(x,y){return y.v-x.v;}).slice(0,10);
  if (top10.length===0) {
    sh.getRange(fila,1).setValue('Sin ventas en el mes actual')
      .setFontStyle('italic').setFontColor('#999999');
  } else {
    var top10Vals = top10.map(function(c,idx){ return [(idx+1)+'. '+c.n, c.v]; });
    sh.getRange(fila, 1, top10Vals.length, 2).setValues(top10Vals);
    top10.forEach(function(c,idx){
      var bg=(idx%2===0)?'#f9f0ff':'#ffffff';
      sh.getRange(fila+idx,1).setBackground(bg);
      sh.getRange(fila+idx,2).setNumberFormat('$#,##0')
        .setHorizontalAlignment('right').setBackground(bg);
    });
    fila += top10.length;
  }

  sh.setFrozenRows(4);
  Logger.log('✅ DASHBOARD escrito — filas: '+fila);
}

// ── HELPERS ───────────────────────────────────────────────────
function _z12_D() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }

function _fechaKey_D(val) {
  if (!val) return null;
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, TZ_DASH, 'yyyy-MM-dd');
    }
    var s = String(val).trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      var d  = parseInt(m[1]);
      var mo = parseInt(m[2]);
      var y  = parseInt(m[3]);
      return y + '-' + (mo < 10 ? '0'+mo : ''+mo) + '-' + (d < 10 ? '0'+d : ''+d);
    }
    return null;
  } catch(e) { return null; }
}

function _keyToDate_D(key) {
  var p = key.split('-');
  return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
}

function _parseFecha_D(val) {
  var key = _fechaKey_D(val);
  if (!key) return null;
  return _keyToDate_D(key);
}

function _normMetodo_D(raw) {
  var k = raw.trim().toUpperCase().replace(/\s+/g,'_');
  if (COMISION_MAP_D[k]) return k;
  return ALIAS_METODO_D[k] || null;
}

function _hdr_D(sh, fila, texto, bg, fg) {
  _mf_D(sh, fila, 1, 1, 15, texto, fg, bg, 10, true);
  return fila + 1;
}

// ← FIX v2.14: _row_D usa setValues batch para los 12 meses + total
function _row_D(sh, fila, label, hoy, meses, fmt, bg, bgHoy, bgTot, mesActual) {
  var total = meses.reduce(function(s,v){return s+v;},0);
  var rowVals = [[label, hoy].concat(meses).concat([total])];
  sh.getRange(fila, 1, 1, 15).setValues(rowVals);

  sh.getRange(fila,1).setBackground(bg).setFontSize(10);
  sh.getRange(fila,2).setNumberFormat(fmt).setHorizontalAlignment('right')
    .setBackground(bgHoy).setFontWeight('bold');
  for (var i=0; i<12; i++) {
    sh.getRange(fila,3+i).setNumberFormat(fmt).setHorizontalAlignment('right')
      .setBackground(i===mesActual?'#e8eaf6':bg);
  }
  sh.getRange(fila,15).setNumberFormat(fmt).setHorizontalAlignment('right')
    .setBackground(bgTot).setFontWeight('bold');
  return fila + 1;
}

function _mf_D(sh, fila, col, rs, cs, txt, fg, bg, sz, bold) {
  sh.getRange(fila,col,rs,cs).merge().setValue(txt)
    .setFontColor(fg).setBackground(bg).setFontSize(sz||10)
    .setFontWeight(bold?'bold':'normal')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
}

function diagnosticarVentas() {
  var ss = SpreadsheetApp.openById(SS_ID_DASH);
  var sh = ss.getSheetByName('VENTAS');
  var rows = sh.getRange(52, 1, 50, 18).getValues();
  rows.forEach(function(r, i) {
    var fechaRaw = r[2];
    var total    = r[8];
    var estado   = r[17];
    var key      = _fechaKey_D(fechaRaw);
    Logger.log('VENTA '+i+': fecha_raw='+fechaRaw
      +' tipo='+typeof fechaRaw
      +' | key='+key
      +' | total='+total
      +' | estado='+estado);
  });
}

// ── QRS: Total $ / Total % / Partida (agregado sobre DASHBOARD) ──
// ← v2.17: filas actualizadas tras agregar Tarjeta Naranja (15 métodos de pago).
// Bloque de forma de pago ahora ocupa filas 12-26 (antes 12-25).
// Comisiones (Total $ / Total neto) se corrieron de 28,29 a 29,30.
var FILAS_QRS_D = [7, 9, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 29, 30];
var FILA_TOTAL_VENTAS_QRS = 7;

function _agregarColumnasQRS_D(ss, anio) {
  var sh = ss.getSheetByName('DASHBOARD');
  if (!sh) { Logger.log('❌ Hoja DASHBOARD no encontrada (QRS)'); return; }

  sh.setColumnWidth(17, 110); // Q
  sh.setColumnWidth(18, 90);  // R
  sh.setColumnWidth(19, 220); // S

  sh.getRange('Q4').setValue('Total $ ' + anio).setFontWeight('bold')
    .setBackground('#1a1a2e').setFontColor('#f0c040').setHorizontalAlignment('center');
  sh.getRange('R4').setValue('Total % ' + anio).setFontWeight('bold')
    .setBackground('#1a1a2e').setFontColor('#f0c040').setHorizontalAlignment('center');
  sh.getRange('S4').setValue('Partida').setFontWeight('bold')
    .setBackground('#1a1a2e').setFontColor('#f0c040').setHorizontalAlignment('center');

  var totalVentasRef = '$O$' + FILA_TOTAL_VENTAS_QRS;

  FILAS_QRS_D.forEach(function(fila) {
    sh.getRange('Q' + fila).setFormula('=O' + fila);
    sh.getRange('Q' + fila).setNumberFormat('$#,##0');

    sh.getRange('R' + fila).setFormula('=IFERROR(Q' + fila + '/' + totalVentasRef + ',0)');
    sh.getRange('R' + fila).setNumberFormat('0.00%');

    sh.getRange('S' + fila).setFormula('=A' + fila);
  });

  Logger.log('✅ Columnas Q/R/S actualizadas en DASHBOARD');
}
// ← NUEVO: lee base_datos_CLI completa para el dropdown/buscador de clientes
function getClientesBase_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const sh = ss.getSheetByName("base_datos_CLI");
  if (!sh || sh.getLastRow() < 2) return { ok: true, clientes: [] };

  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  const clientes = datos
    .filter(fila => fila[0]) // solo filas con CLIENTE_ID
    .map(fila => ({
      clienteId: fila[0],
      nombre: fila[1] || "",
      tel: fila[2] || "",
      email: fila[3] || "",
      cuit: fila[4] || "",
      razonSocial: fila[5] || "",
    }));

  return { ok: true, clientes };
}
// ← NUEVO: actualiza tel/email/cuit/razonSocial de un cliente existente por su CLIENTE_ID
// El Nombre nunca se edita acá — es la clave que usa el proceso diario de matching.
function actualizarClienteBase_V2(data) {
  const clienteId = (data.clienteId || "").trim();
  if (!clienteId) return { ok: false, error: "clienteId requerido" };

  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const sh = ss.getSheetByName("base_datos_CLI");
  if (!sh || sh.getLastRow() < 2) return { ok: false, error: "base_datos_CLI vacía o no encontrada" };

  const datos = sh.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][0] || "").toString().trim() === clienteId) {
      const row = i + 1;
      sh.getRange(row, 3).setValue(data.tel || "");          // C — TEL
      sh.getRange(row, 4).setValue(data.email || "");        // D — EMAIL
      sh.getRange(row, 5).setValue(data.cuit || "");         // E — CUIT
      sh.getRange(row, 6).setValue(data.razonSocial || "");  // F — RAZON_SOCIAL
      return { ok: true, clienteId, actualizado: true };
    }
  }
  return { ok: false, error: "Cliente no encontrado: " + clienteId };
}
