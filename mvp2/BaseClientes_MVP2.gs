// ═══════════════════════════════════════════════════════════
// BASE_DATOS_CLI + asignación de CLIENTE_ID en VENTAS
// Trigger propio, corre 1 vez al día (recomendado: junto con el de Dashboard)
// Escribe SOLO columna U de VENTAS (nunca A:T) y reescribe base_datos_CLI
// combinando datos existentes + nuevos, sin perder ediciones manuales.
// ═══════════════════════════════════════════════════════════

var EXCLUIR_NOMBRES_CLI = ['CONSUMIDOR FINAL', 'CLIENTE MOSTRADOR', ''];
var COL_CLIENTE_ID_VENTAS = 21; // columna U

function triggerBaseClientes() {
  try {
    actualizarBaseYAsignarIDs_D();
  } catch(e) {
    Logger.log('❌ Error en actualizarBaseYAsignarIDs_D: ' + e.message + ' | Stack: ' + e.stack);
  }
}

function actualizarBaseYAsignarIDs_D() {
  var ss = SpreadsheetApp.openById(SS_ID_DASH);
  var shVentas = ss.getSheetByName('VENTAS');
  if (!shVentas || shVentas.getLastRow() < 2) { Logger.log('❌ VENTAS vacía o no encontrada'); return; }

  var shBase = ss.getSheetByName('base_datos_CLI');
  if (!shBase) {
    shBase = ss.insertSheet('base_datos_CLI');
    Logger.log('ℹ️ Hoja base_datos_CLI creada automáticamente');
  }

  var lastRowV = shVentas.getLastRow();
  var ventasData = shVentas.getRange(2, 1, lastRowV - 1, 21).getValues();

  // ── Cargar base existente en memoria (preserva ediciones manuales) ──
  var baseByName = {};
  var maxIdNum = 0;
  var lastRowB = shBase.getLastRow();

  if (lastRowB >= 2) {
    var baseData = shBase.getRange(2, 1, lastRowB - 1, 8).getValues();
    baseData.forEach(function(r) {
      var id = String(r[0] || '').trim();
      var nombre = String(r[1] || '').trim();
      if (!nombre) return;
      var nombreNorm = nombre.toUpperCase();
      var m = id.match(/CLI-(\d+)/);
      if (m) maxIdNum = Math.max(maxIdNum, parseInt(m[1]));
      baseByName[nombreNorm] = {
        id: id,
        nombre: nombre,
        tel: String(r[2] || '').trim(),
        email: String(r[3] || '').trim(),
        cuit: String(r[4] || '').trim(),
        razon: String(r[5] || '').trim(),
        fechaAlta: r[6] || '',
        revisar: String(r[7] || '').trim()
      };
    });
  } else {
    shBase.getRange(1, 1, 1, 8).setValues([[
      'CLIENTE_ID', 'NOMBRE_CLIENTE', 'TEL_CLIENTE', 'EMAIL_CLIENTE',
      'CUIT', 'RAZON_SOCIAL', 'FECHA_ALTA', 'REVISAR_DUPLICADO'
    ]]);
    shBase.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  }

  // ── Procesar VENTAS: asignar ID a filas nuevas, completar datos vacíos ──
  var idsParaVentas = [];
  var huboNuevos = false;

  ventasData.forEach(function(r) {
    var idExistente = String(r[20] || '').trim(); // columna U (índice 20, base 0)
    if (idExistente) {
      idsParaVentas.push([idExistente]);
      return;
    }

    var nombreRaw = String(r[5] || '').trim(); // F
    var nombreNorm = nombreRaw.toUpperCase();

    if (!nombreRaw || EXCLUIR_NOMBRES_CLI.indexOf(nombreNorm) !== -1) {
      idsParaVentas.push(['']);
      return;
    }

    var tel = String(r[6] || '').trim();   // G
    var email = String(r[7] || '').trim(); // H
    var cuit = String(r[12] || '').trim(); // M
    var razon = String(r[13] || '').trim();// N

    // ← FIX: normalizar fecha con los mismos helpers que usa el Dashboard
    var fechaRaw = r[2]; // C
    var fechaKey = _fechaKey_D(fechaRaw);
    var fecha = fechaKey
      ? Utilities.formatDate(_keyToDate_D(fechaKey), TZ_DASH, 'dd/MM/yyyy')
      : fechaRaw;

    var entry = baseByName[nombreNorm];

    if (entry) {
      if (tel && entry.tel && tel !== entry.tel) {
        entry.revisar = 'SI';
      }
      if (!entry.tel && tel) entry.tel = tel;
      if (!entry.email && email) entry.email = email;
      if (!entry.cuit && cuit) entry.cuit = cuit;
      if (!entry.razon && razon) entry.razon = razon;
      idsParaVentas.push([entry.id]);
    } else {
      maxIdNum++;
      var nuevoId = 'CLI-' + ('0000' + maxIdNum).slice(-4);
      var nuevoEntry = {
        id: nuevoId, nombre: nombreRaw, tel: tel, email: email,
        cuit: cuit, razon: razon, fechaAlta: fecha, revisar: ''
      };
      baseByName[nombreNorm] = nuevoEntry;
      idsParaVentas.push([nuevoId]);
      huboNuevos = true;
    }
  });

  // ── Escribir SOLO columna U de VENTAS (nunca toca A:T) ──
  shVentas.getRange(2, COL_CLIENTE_ID_VENTAS, idsParaVentas.length, 1).setValues(idsParaVentas);

  // ── Reescribir base_datos_CLI completa (merge en memoria ya aplicado) ──
  var filasFinal = Object.keys(baseByName).map(function(k) { return baseByName[k]; });
  filasFinal.sort(function(a, b) { return a.id.localeCompare(b.id); });

  if (filasFinal.length > 0) {
    var outVals = filasFinal.map(function(e) {
      return [e.id, e.nombre, e.tel, e.email, e.cuit, e.razon, e.fechaAlta, e.revisar];
    });
    shBase.getRange(2, 1, outVals.length, 8).setValues(outVals);
  }

  shBase.setColumnWidth(1, 90);
  shBase.setColumnWidth(2, 220);
  shBase.setColumnWidth(3, 120);
  shBase.setColumnWidth(4, 200);
  shBase.setColumnWidth(8, 130);
  shBase.setFrozenRows(1);

  Logger.log('✅ base_datos_CLI actualizada — ' + filasFinal.length + ' clientes | Nuevos hoy: ' + huboNuevos);
}

// ═══════════════════════════════════════════════════════════
// FIX PUNTUAL — corre 1 sola vez, normaliza FECHA_ALTA de las
// filas ya existentes en base_datos_CLI a formato dd/MM/yyyy.
// No toca CLIENTE_ID, nombre, tel, email, cuit, razón ni revisar.
// ═══════════════════════════════════════════════════════════
function corregirFechasBaseClientes_D() {
  var ss = SpreadsheetApp.openById(SS_ID_DASH);
  var shBase = ss.getSheetByName('base_datos_CLI');
  if (!shBase || shBase.getLastRow() < 2) { Logger.log('❌ base_datos_CLI vacía o no encontrada'); return; }

  var lastRow = shBase.getLastRow();
  var fechas = shBase.getRange(2, 7, lastRow - 1, 1).getValues(); // columna G = FECHA_ALTA
  var corregidas = 0;

  var fechasNuevas = fechas.map(function(r) {
    var val = r[0];
    if (!val) return [val];
    var key = _fechaKey_D(val);
    if (!key) return [val]; // no se pudo parsear, se deja como está
    var formateada = Utilities.formatDate(_keyToDate_D(key), TZ_DASH, 'dd/MM/yyyy');
    if (formateada !== String(val)) corregidas++;
    return [formateada];
  });

  shBase.getRange(2, 7, fechasNuevas.length, 1).setValues(fechasNuevas);
  Logger.log('✅ Fechas normalizadas en base_datos_CLI — filas corregidas: ' + corregidas);
}
