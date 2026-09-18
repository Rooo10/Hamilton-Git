// ============================================================
// CierrePanel_MVP2.gs  (v2.22)
// Backend del panel "Cierre del día" de la app (Index_MVP2.html).
//
// cierreDeCajaHTML_V2()    → calcula el cierre de HOY, refresca VENTAS_DIARIAS, arma los
//                            3 bloques del panel y guarda una "foto" con fecha/hora del cálculo.
// getUltimoCierreHTML_V2() → devuelve esa foto SIN recalcular y SIN escribir en el Sheet.
//
// Usa variables/funciones de CierreDiario_MVP2.gs y Dashboard_MVP2.gs (mismo proyecto GAS).
// ============================================================

var CLAVE_FOTO_CP = "CIERRE_FOTO";
var TROZO_PROP_CP = 3000; // cada propiedad admite ~9 KB: la foto se guarda en trozos chicos

function cierreDeCajaHTML_V2() {
  try {
    var resultado = calcularCierreDiario();   // 1) cierra HOY: columna del día + ítems en CIERRE_DIARIO
    completarVentasDiarias_VD();              // 2) refresca VENTAS_DIARIAS con los pagos actuales
    SpreadsheetApp.flush();

    var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
    var p = String(resultado.fecha).split("/");        // dd/MM/yyyy → yyyy-MM-dd
    var fechaKey = p[2] + "-" + p[1] + "-" + p[0];

    var foto = {
      fecha: resultado.fecha,
      calculadoEn: Utilities.formatDate(new Date(), TZ_CIERRE, "dd/MM/yyyy HH:mm"),
      indicadores: _indicadoresDelDia_CP(ss, fechaKey),
      items: _itemsDelDia_CP(ss, resultado.itemsCount),
      pagos: _pagosDelDia_CP(ss, fechaKey)
    };
    _guardarJson_CP(CLAVE_FOTO_CP, foto);
    return { ok: true, existe: true, cierre: foto };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getUltimoCierreHTML_V2() {
  try {
    var foto = _leerJson_CP(CLAVE_FOTO_CP);
    if (!foto) return { ok: true, existe: false };
    return { ok: true, existe: true, cierre: foto };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── Bloque 1: indicadores (CIERRE_DIARIO, filas 4 a 17: columna A + columna de la fecha) ──
function _indicadoresDelDia_CP(ss, fechaKey) {
  var sh = ss.getSheetByName(SHEET_CIERRE);
  if (!sh) throw new Error("No existe la hoja " + SHEET_CIERRE);
  var headers = sh.getRange(FILA_HEADER_FECHAS_CD, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = 0;
  for (var c = 1; c < headers.length; c++) {
    if (headers[c] !== "" && _fechaKey_D(headers[c]) === fechaKey) { col = c + 1; break; }
  }
  if (!col) throw new Error("No encontré la columna del " + fechaKey + " en " + SHEET_CIERRE);

  var n = LABELS_INDICADORES_CD.length;
  var etiquetas = sh.getRange(FILA_PRIMER_INDICADOR_CD, 1, n, 1).getValues();
  var valores = sh.getRange(FILA_PRIMER_INDICADOR_CD, col, n, 1).getValues();
  var lista = [];
  for (var i = 0; i < n; i++) {
    var fmt = FORMATO_INDICADORES_CD[LABELS_INDICADORES_CD[i]] || "";
    var tipo = fmt.indexOf("%") >= 0 ? "porcentaje" : (fmt.indexOf("$") >= 0 ? "moneda" : "cantidad");
    var v = valores[i][0];
    var num = (v === "" || v === null || isNaN(Number(v))) ? null : Number(v);
    lista.push({ etiqueta: String(etiquetas[i][0] || LABELS_INDICADORES_CD[i]), valor: num, tipo: tipo });
  }
  return lista;
}

// ── Bloque 2: ítems vendidos (CIERRE_DIARIO, tabla desde FILA_PRIMER_ITEM_CD, columnas B a E) ──
function _itemsDelDia_CP(ss, cantidad) {
  if (!cantidad || cantidad < 1) return [];
  var sh = ss.getSheetByName(SHEET_CIERRE);
  var filas = sh.getRange(FILA_PRIMER_ITEM_CD, 2, cantidad, 4).getValues();
  return filas.map(function(r) {
    return {
      cliente: String(r[0] || ""),
      descripcion: String(r[1] || ""),
      cantidad: Number(r[2]) || 0,
      subtotal: Number(r[3]) || 0
    };
  });
}

// ── Bloque 3: pagos por forma de pago del día (VENTAS_DIARIAS, Tabla Diaria) ──
function _pagosDelDia_CP(ss, fechaKey) {
  var sh = ss.getSheetByName("VENTAS_DIARIAS");
  if (!sh) throw new Error("No existe la hoja VENTAS_DIARIAS");
  var pos = _ubicarTablas_VD(sh);
  if (!pos) throw new Error("No encontré la Tabla Diaria en VENTAS_DIARIAS");

  var nM = METODOS_DASH.length;
  var etiquetas = sh.getRange(pos.filaDiaria - 1, 3, 1, nM).getValues()[0];
  var n = sh.getLastRow() - pos.filaDiaria + 1;
  var fechas = sh.getRange(pos.filaDiaria, 1, n, 1).getValues();
  var fila = 0;
  for (var i = 0; i < n; i++) {
    if (_fechaKey_D(fechas[i][0]) === fechaKey) { fila = pos.filaDiaria + i; break; }
  }
  if (!fila) return { metodos: [], total: 0, aviso: "No encontré la fila de esa fecha en VENTAS_DIARIAS" };

  var valores = sh.getRange(fila, 3, 1, nM + 1).getValues()[0]; // columnas C..Q + R (Total)
  var metodos = [], suma = 0;
  for (var j = 0; j < nM; j++) {
    var monto = Number(valores[j]) || 0;
    suma += monto;
    metodos.push({ etiqueta: String(etiquetas[j] || METODOS_DASH[j].label), monto: monto });
  }
  var totalHoja = valores[nM];
  return { metodos: metodos, total: (typeof totalHoja === "number") ? totalHoja : suma };
}

// ── Guardado de la foto en Script Properties (en trozos, por el límite de tamaño) ──
function _guardarJson_CP(clave, obj) {
  var props = PropertiesService.getScriptProperties();
  var texto = JSON.stringify(obj);
  var nuevos = Math.ceil(texto.length / TROZO_PROP_CP);
  var previos = parseInt(props.getProperty(clave + "_N"), 10) || 0;
  var lote = {};
  for (var i = 0; i < nuevos; i++) lote[clave + "_" + i] = texto.substr(i * TROZO_PROP_CP, TROZO_PROP_CP);
  lote[clave + "_N"] = String(nuevos);
  props.setProperties(lote);
  for (var k = nuevos; k < previos; k++) props.deleteProperty(clave + "_" + k); // borra trozos sobrantes
}

function _leerJson_CP(clave) {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty(clave + "_N"), 10) || 0;
  if (n < 1) return null;
  var texto = "";
  for (var i = 0; i < n; i++) {
    var t = props.getProperty(clave + "_" + i);
    if (t === null) return null; // foto incompleta: se ignora
    texto += t;
  }
  return JSON.parse(texto);
}

// ── TESTING (correr desde el editor): muestra en el Registro cómo quedó el cierre de HOY ──
function test_cierreDeCajaHTML_V2() {
  var r = cierreDeCajaHTML_V2();
  Logger.log(JSON.stringify(r, null, 1).substring(0, 4000));
}

// ← v2.24: recalcula el cierre de días PASADOS (usa los datos que ya están en VENTAS, PAGOS,
// ITEMS, etc.). No toca el bloque "Detalle de ítems" (igual que el backfill) ni ninguna venta:
// solo vuelve a escribir la columna de cada día en CIERRE_DIARIO.
// fechasKey: lista de fechas como 'yyyy-MM-dd'. Ojo con el orden: una fecha que todavía no
// tiene columna se agrega al final de la hoja.
function recalcularCierresDias_CD(fechasKey) {
  var ss = SpreadsheetApp.openById(SS_ID_CIERRE);
  var sh = _asegurarHojaCierre_CD(ss);
  var ventas   = _leerHoja_CD(ss, 'VENTAS', 18);
  var pagos    = _leerHoja_CD(ss, 'PAGOS', 8);
  var pagosCC  = _leerHoja_CD(ss, 'PAGOS_CC', 4);
  var facturas = _leerHoja_CD(ss, 'FACTURAS_PENDIENTES', 15);
  var items    = _leerHoja_CD(ss, 'ITEMS', 8);
  var comMap   = _leerComisiones_D(ss);
  fechasKey.forEach(function(fechaKey) {
    var p = fechaKey.split('-');
    var fechaObj = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    var r = _procesarUnDia_CD(sh, fechaObj, fechaKey, ventas, pagos, pagosCC, facturas, items, comMap, false);
    Logger.log('✅ ' + r.fecha + ' recalculado | Ventas: $' + r.totalVentas + ' (' + r.cantVentas +
               ') | Pagos por ventas: $' + r.totalPagosVentas + ' | Pagos contra CC: $' + r.totalPagosCC);
  });
  SpreadsheetApp.flush();
}

// Correr desde el editor: recalcula 16/09 y 17/09/2026 (en este orden).
function recalcularCierres_16y17_sept() {
  recalcularCierresDias_CD(['2026-09-16', '2026-09-17']);
}
