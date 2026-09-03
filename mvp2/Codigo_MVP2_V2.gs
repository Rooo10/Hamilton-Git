const CONFIG_V2 = {
  SPREADSHEET_ID:      "18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8",
  EMAIL_ENCARGADO:     "HamiltonDecoRosario@gmail.com",
  NOMBRE_LOCAL:        "Hamilton Deco",
  DIRECCION_LOCAL:     "Necochea 1427, Rosario",
  TELEFONO_LOCAL:      "(341) 255-6695",
  PLAZO_PAGO_CC_DIAS:  30,
  SPREADSHEET_URL:     "https://docs.google.com/spreadsheets/d/18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8/edit",
  TZ:                  "America/Argentina/Buenos_Aires",
};

// ← FIX v2.15: helper central — escribe fechas como string DD/MM/YYYY
function _fechaStr_V2(date) {
  return Utilities.formatDate(date, CONFIG_V2.TZ, "dd/MM/yyyy");
}

// ← FIX v2.16: fuerza una celda a texto plano y re-escribe el valor,
// para blindarla contra la reinterpretación automática de Sheets
// (que puede invertir DD/MM cuando el día es ≤12).
function _forzarFechaTexto_V2(sheet, row, col, valorStr) {
  sheet.getRange(row, col).setNumberFormat('@').setValue(valorStr);
}

// ← NUEVO v2.18: resuelve la fecha de venta elegida por el operador (opcional).
// Valida formato estricto YYYY-MM-DD y que sea una fecha real (rechaza ej. 2026-02-30).
// Si falta o es inválida, hace fallback SILENCIOSO a "now" — nunca bloquea el registro.
function _resolverFechaVenta_V2(fechaVentaRaw, now) {
  if (typeof fechaVentaRaw === "string" && /^\d{4}-\d{2}-\d{2}$/.test(fechaVentaRaw)) {
    const partes = fechaVentaRaw.split("-");
    const anio = parseInt(partes[0], 10);
    const mes  = parseInt(partes[1], 10);
    const dia  = parseInt(partes[2], 10);
    const fechaDate = new Date(anio, mes - 1, dia, now.getHours(), now.getMinutes(), now.getSeconds());
    if (fechaDate.getFullYear() === anio && (fechaDate.getMonth() + 1) === mes && fechaDate.getDate() === dia) {
      return { fechaStr: _fechaStr_V2(fechaDate), fechaDate: fechaDate };
    }
  }
  return { fechaStr: _fechaStr_V2(now), fechaDate: now };
}

function registrarVenta_V2(data) {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const now = new Date();
  // ← NUEVO v2.18: fecha de venta editable (opcional). Fallback silencioso a "now" si viene vacía o inválida.
  const fechaResuelta = _resolverFechaVenta_V2(data.fechaVenta, now);
  const fechaStr = fechaResuelta.fechaStr;
  const fechaVentaDate = fechaResuelta.fechaDate;
  const idVenta = _generarId_V2("V");
  const idRecibo = _generarId_V2("R");

  let pagosDetalle = [];
  try { pagosDetalle = JSON.parse(data.pagosDetalle || "[]"); } catch(e) { pagosDetalle = []; }
  let formaPagoLabel = "";
  if (pagosDetalle.length === 1) {
    formaPagoLabel = pagosDetalle[0].metodo || "EFECTIVO";
  } else if (pagosDetalle.length > 1) {
    formaPagoLabel = "MIXTO: " + pagosDetalle.map(p => p.metodo).join("+");
  } else {
    formaPagoLabel = data.formaPago || "EFECTIVO";
  }

  const hora      = now.toLocaleTimeString("es-AR", {hour:"2-digit", minute:"2-digit", second:"2-digit"});
  const timestamp = now.toISOString();

  const shVentas = ss.getSheetByName("VENTAS");
  shVentas.appendRow([
    idVenta,                          // A — ID_VENTA
    idRecibo,                         // B — ID_RECIBO
    fechaStr,                         // C — FECHA
    hora,                             // D — HORA
    timestamp,                        // E — TIMESTAMP
    data.nombreCliente || "",         // F — NOMBRE_CLIENTE
    data.telCliente || "",            // G — TEL_CLIENTE
    data.emailCliente || "",          // H — EMAIL_CLIENTE
    Number(data.total) || 0,          // I — TOTAL
    formaPagoLabel,                   // J — FORMA_PAGO
    data.factura || "NO",             // K — FACTURA
    data.tipoEnvioFact || "",         // L — TIPO_ENVIO_FACT
    data.cuit || "",                  // M — CUIT
    data.razonSocial || "",           // N — RAZON_SOCIAL
    Number(data.importeFactura) || 0, // O — IMPORTE_FACTURA
    data.importeFacturaTipo || "",    // P — IMPORTE_FACTURA_TIPO
    data.envioRecibo || "NINGUNO",    // Q — ENVIO_RECIBO
    "REGISTRADA",                     // R — ESTADO
    data.notas || "",                 // S — NOTAS
    "",                               // T — NRO_FACTURA
  ]);
  // ← FIX v2.16: blindar columna C (FECHA) contra reinterpretación de Sheets
  _forzarFechaTexto_V2(shVentas, shVentas.getLastRow(), 3, fechaStr);

  const shItems = ss.getSheetByName("ITEMS");
  (data.items || []).forEach(item => {
    shItems.appendRow([
      idVenta, fechaStr, data.nombreCliente || "",
      item.descripcion || "", Number(item.cantidad) || 1,
      Number(item.precioUnitario) || 0, Number(item.subtotal) || 0, "REGISTRADA",
    ]);
    // ← FIX v2.16: columna B (FECHA) de ITEMS
    _forzarFechaTexto_V2(shItems, shItems.getLastRow(), 2, fechaStr);
  });

  const shPagos = ss.getSheetByName("PAGOS");
  if (pagosDetalle.length > 0) {
    pagosDetalle.forEach(p => {
      const monto = Number(p.importe) || 0;
      if (monto > 0) {
        shPagos.appendRow([
          _generarId_V2("P"), idVenta, fechaStr, data.nombreCliente || "",
          monto, p.metodo || "EFECTIVO", p.cuotas || "1", "REGISTRADA",
        ]);
        // ← FIX v2.16: columna C (FECHA) de PAGOS
        _forzarFechaTexto_V2(shPagos, shPagos.getLastRow(), 3, fechaStr);
      }
    });
  } else {
    const montoEfec = Number(data.montoEfectivo) || 0;
    const montoTarj = Number(data.montoTarjeta) || 0;
    const montoCC   = Number(data.montoCC) || 0;
    if (montoEfec > 0) {
      shPagos.appendRow([_generarId_V2("P"), idVenta, fechaStr, data.nombreCliente || "", montoEfec, "EFECTIVO", "1", "REGISTRADA"]);
      _forzarFechaTexto_V2(shPagos, shPagos.getLastRow(), 3, fechaStr);
    }
    if (montoTarj > 0) {
      shPagos.appendRow([_generarId_V2("P"), idVenta, fechaStr, data.nombreCliente || "", montoTarj, "TARJETA", "1", "REGISTRADA"]);
      _forzarFechaTexto_V2(shPagos, shPagos.getLastRow(), 3, fechaStr);
    }
    if (montoCC > 0) {
      shPagos.appendRow([_generarId_V2("P"), idVenta, fechaStr, data.nombreCliente || "", montoCC, "CUENTA_CORRIENTE", "1", "REGISTRADA"]);
      _forzarFechaTexto_V2(shPagos, shPagos.getLastRow(), 3, fechaStr);
    }
  }

  let idCC = null;
  const montoCC = Number(data.montoCC) || 0;
  const formaPagoOriginal = data.formaPago || "EFECTIVO";
  if ((formaPagoOriginal === "CC" || formaPagoOriginal === "MIXTO") && montoCC > 0) {
    idCC = _generarId_V2("CC");
    // ← v2.18: el vencimiento se calcula desde la fecha de venta elegida, no desde "ahora"
    const fechaVenc = new Date(fechaVentaDate.getTime() + CONFIG_V2.PLAZO_PAGO_CC_DIAS * 86400000);
    const fechaVencStr = _fechaStr_V2(fechaVenc);
    const shCCNueva = ss.getSheetByName("CC");
    shCCNueva.appendRow([
      idCC, idVenta, fechaStr, data.nombreCliente || "", data.telCliente || "", data.emailCliente || "",
      montoCC, 0, fechaVencStr, "", 0, "SIN_VENCER",
    ]);
    // ← FIX v2.16: columna C (fecha compra) y columna I (fecha vencimiento) de CC
    const rowCC = shCCNueva.getLastRow();
    _forzarFechaTexto_V2(shCCNueva, rowCC, 3, fechaStr);
    _forzarFechaTexto_V2(shCCNueva, rowCC, 9, fechaVencStr);
  }

  if ((data.factura || "NO") === "SI") {
    const itemsResumen = (data.items || []).map(i => i.cantidad + "x " + i.descripcion).join(", ");
    const shFactNueva = ss.getSheetByName("FACTURAS_PENDIENTES");
    shFactNueva.appendRow([
      idVenta, fechaStr, data.nombreCliente || "", data.telCliente || "", data.emailCliente || "",
      Number(data.total) || 0, formaPagoLabel, data.tipoEnvioFact || "",
      data.cuit || "", data.razonSocial || "",
      "", "", "", "",
      "PENDIENTE",
      itemsResumen, data.notas || "",
    ]);
    // ← FIX v2.16: columna B (FECHA) de FACTURAS_PENDIENTES
    _forzarFechaTexto_V2(shFactNueva, shFactNueva.getLastRow(), 2, fechaStr);

    const asunto = "ACCION: Emitir + Enviar Factura -> " + (data.nombreCliente || "") + " | $" + Number(data.total) + " | Recibo N " + idRecibo;
    const cuerpo = "<p><strong>Cliente:</strong> " + (data.nombreCliente || "") + "</p>" +
      "<p><strong>CUIT/DNI:</strong> " + (data.cuit || "") + "</p>" +
      "<p><strong>Razon Social:</strong> " + (data.razonSocial || "") + "</p>" +
      "<p><strong>Email cliente:</strong> " + (data.emailCliente || "") + "</p>" +
      "<p><strong>Tel cliente:</strong> " + (data.telCliente || "") + "</p>" +
      "<p><strong>Importe a facturar:</strong> $" + Number(data.importeFactura) + " (" + (data.importeFacturaTipo || "") + ")</p>" +
      "<p><strong>Envio factura:</strong> " + (data.tipoEnvioFact || "") + "</p>" +
      "<p><strong>ID Venta:</strong> " + idVenta + "</p>";
    _enviarEmailAccion_V2(asunto, cuerpo);
  }

  if (formaPagoOriginal === "CC" || formaPagoOriginal === "MIXTO" ||
      formaPagoLabel.includes("CUENTA_CORRIENTE")) {
    _reconstruirListadoCC_V2(ss);
  }

  _enviarRecibo_V2(data, idVenta, idRecibo, now, formaPagoLabel);
  return { ok: true, idVenta, idRecibo, idCC };
}

function anularVenta_V2(data) {
  const idVenta = data.idVenta;
  if (!idVenta) return { ok: false, error: "idVenta requerido" };
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  _actualizarEstadoPorId_V2(ss.getSheetByName("VENTAS"), 1, idVenta, 18, "ANULADA");
  _actualizarEstadoPorIdMultiple_V2(ss.getSheetByName("ITEMS"), 1, idVenta, 8, "ANULADA");
  _actualizarEstadoPorIdMultiple_V2(ss.getSheetByName("PAGOS"), 2, idVenta, 8, "ANULADA");
  _actualizarEstadoPorId_V2(ss.getSheetByName("CC"), 2, idVenta, 12, "ANULADA");
  _actualizarEstadoPorId_V2(ss.getSheetByName("FACTURAS_PENDIENTES"), 1, idVenta, 15, "ANULADA");
  return { ok: true, idVenta, anulada: true };
}

function registrarPagoCC_V2(data) {
  const nombreCliente = (data.nombreCliente || "").trim();
  const montoPagado = Number(data.montoPagado) || 0;
  if (!nombreCliente) return { ok: false, error: "nombreCliente requerido" };
  if (montoPagado <= 0) return { ok: false, error: "montoPagado debe ser mayor a 0" };

  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const shCC = ss.getSheetByName("CC");
  const now = new Date();
  const fechaStr = _fechaStr_V2(now);

  const shPagosCCNueva = ss.getSheetByName("PAGOS_CC");
  shPagosCCNueva.appendRow([_generarId_V2("PCC"), nombreCliente, fechaStr, montoPagado]);
  // ← FIX v2.16: columna C (FECHA) de PAGOS_CC
  _forzarFechaTexto_V2(shPagosCCNueva, shPagosCCNueva.getLastRow(), 3, fechaStr);

  const datos = shCC.getDataRange().getValues();
  const filasCliente = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const nombre = (fila[3] || "").trim().toLowerCase();
    const estado = (fila[11] || "").toString().trim();
    if (nombre !== nombreCliente.toLowerCase()) continue;
    if (estado !== "SIN_VENCER" && estado !== "VENCIDO") continue;
    const importe = Number(fila[6]) || 0;
    const pagosParciales = Number(fila[7]) || 0;
    const saldoPendiente = importe - pagosParciales;
    if (saldoPendiente > 0) {
      filasCliente.push({
        rowIndex: i + 1, idCC: fila[0], fechaCompra: fila[2],
        importe, pagosParciales, saldoPendiente,
        diasMora: Number(fila[10]) || 0, estado,
      });
    }
  }

  if (filasCliente.length === 0)
    return { ok: false, error: "No hay deudas activas para " + nombreCliente };

  filasCliente.sort((a, b) => b.diasMora !== a.diasMora
    ? b.diasMora - a.diasMora
    : new Date(a.fechaCompra) - new Date(b.fechaCompra));

  let saldoDisponible = montoPagado;
  const detalle = [];

  for (const fila of filasCliente) {
    if (saldoDisponible <= 0) break;
    const aImputar = Math.min(saldoDisponible, fila.saldoPendiente);
    const nuevosPagosParciales = fila.pagosParciales + aImputar;
    saldoDisponible -= aImputar;
    const saldoRestante = fila.importe - nuevosPagosParciales;
    const nuevoEstado = saldoRestante <= 0.001 ? "PAGADO" : fila.estado;
    shCC.getRange(fila.rowIndex, 8).setValue(nuevosPagosParciales);
    // ← FIX v2.16: columna J (fecha último pago) de CC
    _forzarFechaTexto_V2(shCC, fila.rowIndex, 10, fechaStr);
    shCC.getRange(fila.rowIndex, 12).setValue(nuevoEstado);
    if (nuevoEstado === "PAGADO") {
      _limpiarFormatoFila_V2(shCC, fila.rowIndex, 12);
      shCC.getRange(fila.rowIndex, 11).setValue(fila.diasMora);
    }
    detalle.push({ idCC: fila.idCC, imputado: aImputar, estado: nuevoEstado });
  }

  _reconstruirListadoCC_V2(ss);
  return { ok: true, montoPagado, saldoExcedente: saldoDisponible, detalle };
}

function actualizarEstadosCC_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const shCC = ss.getSheetByName("CC");
  const datos = shCC.getDataRange().getValues();
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const estado = (fila[11] || "").toString().trim();
    if (estado !== "SIN_VENCER" && estado !== "VENCIDO") continue;
    const fechaVenc = new Date(fila[8]); fechaVenc.setHours(0,0,0,0);
    const diffDias = Math.floor((hoy - fechaVenc) / 86400000);
    const row = i + 1;
    if (diffDias <= 0) {
      shCC.getRange(row, 11).setValue(0);
      shCC.getRange(row, 12).setValue("SIN_VENCER");
      _limpiarFormatoFila_V2(shCC, row, 12);
    } else {
      shCC.getRange(row, 11).setValue(diffDias);
      shCC.getRange(row, 12).setValue("VENCIDO");
      _aplicarFormatoVencido_V2(shCC, row, 12);
    }
  }
  _reconstruirListadoCC_V2(ss);
}

function onFacturaPendienteEdit_V2(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "FACTURAS_PENDIENTES") return;
    if (e.range.getColumn() !== 11) return;
    const nroFactura = (e.range.getValue() || "").toString().trim();
    if (!nroFactura) return;
    const row = e.range.getRow();
    if (row <= 1) return;
    const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
    const shFact = ss.getSheetByName("FACTURAS_PENDIENTES");
    const fila = shFact.getRange(row, 1, 1, 17).getValues()[0];
    const idVenta       = (fila[0]  || "").toString().trim();
    const nombreCliente = (fila[2]  || "").toString().trim();
    const telCliente    = (fila[3]  || "").toString().trim();
    const emailCliente  = (fila[4]  || "").toString().trim();
    const total         = Number(fila[5]) || 0;
    const tipoEnvio     = (fila[7]  || "").toString().trim().toUpperCase();
    const pdfUrl        = (fila[12] || "").toString().trim();
    const now           = new Date();
    const fechaStr      = _fechaStr_V2(now);
    // ← FIX v2.16: columna L (fecha emisión) de FACTURAS_PENDIENTES
    _forzarFechaTexto_V2(shFact, row, 12, fechaStr);
    shFact.getRange(row, 15).setValue("EMITIDA");
    if (idVenta) {
      const shVentas = ss.getSheetByName("VENTAS");
      const datosVentas = shVentas.getDataRange().getValues();
      for (let i = 1; i < datosVentas.length; i++) {
        if ((datosVentas[i][0] || "").toString() === idVenta) {
          shVentas.getRange(i + 1, 18).setValue("COMPLETADA");
          shVentas.getRange(i + 1, 20).setValue(nroFactura);
          break;
        }
      }
    }
    if ((tipoEnvio === "EMAIL" || tipoEnvio === "AMBOS") && emailCliente) {
      try {
        MailApp.sendEmail(emailCliente, "Factura " + nroFactura + " - " + CONFIG_V2.NOMBRE_LOCAL, "",
          { htmlBody: _htmlFactura_V2(nombreCliente, nroFactura, total, pdfUrl, now),
            name: CONFIG_V2.NOMBRE_LOCAL, bcc: CONFIG_V2.EMAIL_ENCARGADO });
      } catch(err) { Logger.log("Error email factura: " + err.message); }
    }
    if ((tipoEnvio === "WHATSAPP" || tipoEnvio === "AMBOS") && telCliente) {
      let msg = "Hola " + nombreCliente + ", su factura N " + nroFactura + " de " + CONFIG_V2.NOMBRE_LOCAL + " ha sido emitida.";
      msg += pdfUrl ? " Puede descargarla aqui: " + pdfUrl : " El PDF sera enviado en breve.";
      shFact.getRange(row, 14).setValue(_armarLinkWhatsApp_V2(telCliente, msg));
    }
  } catch(err) { Logger.log("Error onFacturaPendienteEdit_V2: " + err.message); }
}

function buscarVentas_V2(q) {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const datos = ss.getSheetByName("VENTAS").getDataRange().getValues();
  const datosItems = ss.getSheetByName("ITEMS").getDataRange().getValues();
  const qLow = (q || "").toLowerCase().trim();
  const resultados = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const hayMatch = !qLow ||
      (fila[0]||"").toString().toLowerCase().includes(qLow) ||
      (fila[5]||"").toString().toLowerCase().includes(qLow) ||
      (fila[9]||"").toString().toLowerCase().includes(qLow) ||
      (fila[18]||"").toString().toLowerCase().includes(qLow);
    if (!hayMatch) continue;
    const idVenta = fila[0];
    const items = datosItems.filter(it => it[0] === idVenta).map(it => ({
      descripcion: it[3], cantidad: it[4], precioUnitario: it[5], subtotal: it[6],
    }));
    resultados.push({
      idVenta: fila[0], idRecibo: fila[1],
      fecha: fila[2], hora: fila[3], timestamp: fila[4],
      nombreCliente: fila[5], telCliente: fila[6], emailCliente: fila[7],
      total: fila[8], formaPago: fila[9],
      factura: fila[10], tipoEnvioFact: fila[11], cuit: fila[12], razonSocial: fila[13],
      importeFactura: fila[14], importeFacturaTipo: fila[15],
      envioRecibo: fila[16], estado: fila[17], notas: fila[18], nroFactura: fila[19],
      items,
    });
  }
  resultados.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return { ok: true, total: resultados.length, ventas: resultados };
}

function getVenta_V2(idVenta) {
  const resultado = buscarVentas_V2(idVenta);
  const venta = (resultado.ventas || []).find(v => v.idVenta === idVenta);
  if (!venta) return { ok: false, error: "Venta no encontrada: " + idVenta };
  return { ok: true, venta };
}

function getClientesCC_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const datos = ss.getSheetByName("CC").getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const nombre = (fila[3] || "").trim();
    const estado = (fila[11] || "").toString().trim();
    if (!nombre || (estado !== "SIN_VENCER" && estado !== "VENCIDO")) continue;
    const saldo = (Number(fila[6]) || 0) - (Number(fila[7]) || 0);
    if (saldo <= 0) continue;
    if (!mapa[nombre]) mapa[nombre] = { nombre, saldoTotal: 0, tieneVencido: false };
    mapa[nombre].saldoTotal += saldo;
    if (estado === "VENCIDO") mapa[nombre].tieneVencido = true;
  }
  return { ok: true, clientes: Object.values(mapa).sort((a,b) => a.nombre.localeCompare(b.nombre)) };
}

function getListadoCC_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  _reconstruirListadoCC_V2(ss);
  const datos = ss.getSheetByName("LISTADO_CC").getDataRange().getValues();
  const listado = [];
  for (let i = 2; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0]) continue;
    listado.push({
      nombreCliente: fila[0], totalCompraCC: fila[1], totalPagadoCC: fila[2],
      saldoCCTotal: fila[3], saldoSinVencer: fila[4], saldoVencido: fila[5],
      diasMoraPromPonderado: fila[6], status: fila[7], linkWhatsApp: fila[8],
    });
  }
  return { ok: true, listado };
}

function _reconstruirListadoCC_V2(ss) {
  const shCC = ss.getSheetByName("CC");
  const shPagosCC = ss.getSheetByName("PAGOS_CC");
  const shListado = ss.getSheetByName("LISTADO_CC");
  const anoActual = new Date().getFullYear();
  const datosCC = shCC.getDataRange().getValues();
  const datosPagosCC = shPagosCC.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < datosCC.length; i++) {
    const fila = datosCC[i];
    const nombre = (fila[3] || "").trim();
    const estado = (fila[11] || "").toString().trim();
    if (!nombre || estado === "ANULADA") continue;
    if (!mapa[nombre]) mapa[nombre] = {
      nombre, totalCompraCC: 0, saldoSinVencer: 0, saldoVencido: 0,
      sumaDiasMora: 0, sumaImporteVencido: 0, tel: fila[4] || "",
    };
    const importe = Number(fila[6]) || 0;
    const pagos = Number(fila[7]) || 0;
    const saldo = importe - pagos;
    if (new Date(fila[2]).getFullYear() === anoActual) mapa[nombre].totalCompraCC += importe;
    if (estado === "SIN_VENCER" && saldo > 0) mapa[nombre].saldoSinVencer += saldo;
    if (estado === "VENCIDO" && saldo > 0) {
      const dias = Number(fila[10]) || 0;
      mapa[nombre].saldoVencido += saldo;
      mapa[nombre].sumaDiasMora += dias * saldo;
      mapa[nombre].sumaImporteVencido += saldo;
    }
  }
  const pagosAno = {};
  for (let i = 1; i < datosPagosCC.length; i++) {
    const fila = datosPagosCC[i];
    const nombre = (fila[1] || "").trim();
    if (new Date(fila[2]).getFullYear() === anoActual)
      pagosAno[nombre] = (pagosAno[nombre] || 0) + (Number(fila[3]) || 0);
  }
  const lastRow = shListado.getLastRow();
  if (lastRow >= 3) shListado.getRange(3, 1, lastRow - 2, 9).clearContent().clearFormat();
  const clientes = Object.values(mapa).filter(c => c.saldoSinVencer + c.saldoVencido > 0 || c.totalCompraCC > 0);
  clientes.sort((a, b) => a.nombre.localeCompare(b.nombre));
  clientes.forEach((c, idx) => {
    const row = idx + 3;
    const totalPagado = pagosAno[c.nombre] || 0;
    const saldoTotal = c.saldoSinVencer + c.saldoVencido;
    const moraPromPond = c.sumaImporteVencido > 0 ? Math.round(c.sumaDiasMora / c.sumaImporteVencido) : 0;
    const status = saldoTotal > 0 ? "CON_DEUDA_ACTIVA" : "PAGADO";
    const msg = c.saldoVencido > 0
      ? "Hola " + c.nombre + ", tiene una deuda vencida de $" + c.saldoVencido.toFixed(2) + " en " + CONFIG_V2.NOMBRE_LOCAL + ". Por favor comuniquese para regularizar."
      : "Hola " + c.nombre + ", tiene un saldo pendiente de $" + saldoTotal.toFixed(2) + " en " + CONFIG_V2.NOMBRE_LOCAL + ". Ante cualquier consulta, estamos a su disposicion.";
    const linkWA = c.tel ? _armarLinkWhatsApp_V2(c.tel, msg) : "";
    shListado.getRange(row, 1, 1, 9).setValues([[
      c.nombre, c.totalCompraCC, totalPagado, saldoTotal,
      c.saldoSinVencer, c.saldoVencido, moraPromPond, status, linkWA,
    ]]);
    if (c.saldoVencido > 0) {
      shListado.getRange(row, 1, 1, 9).setBackground("#fce8e6");
      shListado.getRange(row, 6).setFontWeight("bold").setFontColor("#c0392b");
    }
  });
}

function _enviarRecibo_V2(data, idVenta, idRecibo, fecha, formaPagoLabel) {
  const envio = (data.envioRecibo || "NINGUNO").toUpperCase();
  if (envio === "NINGUNO") return;
  const cuerpoHtml = _htmlRecibo_V2(data, idVenta, idRecibo, fecha, formaPagoLabel);
  if ((envio === "EMAIL" || envio === "AMBOS") && data.emailCliente) {
    try {
      MailApp.sendEmail(data.emailCliente, "Recibo de compra - " + CONFIG_V2.NOMBRE_LOCAL, "",
        { htmlBody: cuerpoHtml, name: CONFIG_V2.NOMBRE_LOCAL, bcc: CONFIG_V2.EMAIL_ENCARGADO });
    } catch(err) { Logger.log("Error email recibo: " + err.message); }
  }
}

function _htmlRecibo_V2(data, idVenta, idRecibo, fecha, formaPagoLabel) {
  const items = (data.items || []).map(i =>
    "<tr><td>" + i.descripcion + "</td><td style='text-align:center'>" + i.cantidad +
    "</td><td style='text-align:right'>$" + Number(i.precioUnitario).toFixed(2) +
    "</td><td style='text-align:right'>$" + Number(i.subtotal).toFixed(2) + "</td></tr>"
  ).join("");
  return "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:24px;border:1px solid #e0e0e0;border-radius:8px;'>" +
    "<h2 style='color:#1a1a2e;margin-bottom:4px;'>" + CONFIG_V2.NOMBRE_LOCAL + "</h2>" +
    "<p style='color:#666;margin-top:0;'>" + CONFIG_V2.DIRECCION_LOCAL + " - " + CONFIG_V2.TELEFONO_LOCAL + "</p>" +
    "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'>" +
    "<p><strong>Recibo N:</strong> " + idRecibo + "</p>" +
    "<p><strong>Fecha:</strong> " + new Date(fecha).toLocaleString("es-AR") + "</p>" +
    "<p><strong>Cliente:</strong> " + (data.nombreCliente || "") + "</p>" +
    "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'>" +
    "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
    "<thead><tr style='background:#1a1a2e;color:#fff;'><th style='padding:8px;text-align:left;'>Descripcion</th>" +
    "<th style='padding:8px;text-align:center;'>Cant.</th><th style='padding:8px;text-align:right;'>P.Unit.</th>" +
    "<th style='padding:8px;text-align:right;'>Subtotal</th></tr></thead>" +
    "<tbody>" + items + "</tbody></table>" +
    "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'>" +
    "<p style='font-size:18px;font-weight:bold;text-align:right;'>TOTAL: $" + Number(data.total).toFixed(2) + "</p>" +
    "<p style='font-size:13px;color:#666;text-align:right;'>Forma de pago: " + (formaPagoLabel || data.formaPago || "") + "</p>" +
    "<p style='font-size:12px;color:#999;text-align:center;'>Gracias por su compra en " + CONFIG_V2.NOMBRE_LOCAL + "</p></div>";
}

function _htmlFactura_V2(nombreCliente, nroFactura, total, pdfUrl, fecha) {
  const linkSeccion = pdfUrl
    ? "<p style='text-align:center;margin:24px 0;'><a href='" + pdfUrl + "' style='display:inline-block;padding:12px 28px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:4px;font-weight:bold;font-size:16px;'>Descargar Factura PDF</a></p>"
    : "<p style='color:#666;font-size:14px;'>El PDF de su factura sera enviado en breve.</p>";
  return "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:24px;border:1px solid #e0e0e0;border-radius:8px;'>" +
    "<h2 style='color:#1a1a2e;margin-bottom:4px;'>" + CONFIG_V2.NOMBRE_LOCAL + "</h2>" +
    "<p style='color:#666;margin-top:0;'>" + CONFIG_V2.DIRECCION_LOCAL + " - " + CONFIG_V2.TELEFONO_LOCAL + "</p>" +
    "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'>" +
    "<h3 style='color:#1a1a2e;'>Factura emitida</h3>" +
    "<p><strong>Cliente:</strong> " + nombreCliente + "</p>" +
    "<p><strong>N Factura:</strong> " + nroFactura + "</p>" +
    "<p><strong>Fecha de emision:</strong> " + new Date(fecha).toLocaleDateString("es-AR") + "</p>" +
    "<p><strong>Total facturado:</strong> $" + Number(total).toFixed(2) + "</p>" +
    "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'>" +
    linkSeccion +
    "<p style='font-size:12px;color:#999;text-align:center;'>Gracias por su compra en " + CONFIG_V2.NOMBRE_LOCAL + "</p></div>";
}

function triggerSemanalCC_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  _reconstruirListadoCC_V2(ss);
  const datos = ss.getSheetByName("LISTADO_CC").getDataRange().getValues();
  let tabla = "<table style='width:100%;border-collapse:collapse;font-size:13px;'><thead><tr style='background:#1a1a2e;color:#fff;'><th style='padding:8px;'>Cliente</th><th style='padding:8px;text-align:right;'>Saldo Total</th><th style='padding:8px;text-align:right;'>Vencido</th><th style='padding:8px;text-align:right;'>Mora (dias)</th><th style='padding:8px;'>Status</th></tr></thead><tbody>";
  let hay = false;
  for (let i = 2; i < datos.length; i++) {
    const f = datos[i];
    if (!f[0]) continue;
    hay = true;
    const v = f[5] > 0;
    tabla += "<tr style='background:" + (v?"#fce8e6":"#fff") + ";'><td style='padding:8px;'>" + f[0] + "</td><td style='padding:8px;text-align:right;'>$" + Number(f[3]).toFixed(2) + "</td><td style='padding:8px;text-align:right;color:" + (v?"#c0392b":"inherit") + ";font-weight:" + (v?"bold":"normal") + ";'>$" + Number(f[5]).toFixed(2) + "</td><td style='padding:8px;text-align:right;'>" + f[6] + "</td><td style='padding:8px;'>" + f[7] + "</td></tr>";
  }
  tabla += "</tbody></table>";
  if (!hay) tabla = "<p>No hay cuentas corrientes activas esta semana.</p>";
  _enviarEmailAccion_V2("Resumen semanal CC - " + CONFIG_V2.NOMBRE_LOCAL,
    "<p>Resumen de Cuentas Corrientes al " + new Date().toLocaleDateString("es-AR") + ":</p>" + tabla);
}

function _enviarEmailAccion_V2(asunto, cuerpo) {
  try {
    MailApp.sendEmail(CONFIG_V2.EMAIL_ENCARGADO, asunto, "",
      { htmlBody: "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:24px;'>" +
        "<h2 style='color:#1a1a2e;'>" + CONFIG_V2.NOMBRE_LOCAL + " - Notificacion</h2>" +
        "<hr style='border:none;border-top:1px solid #e0e0e0;'>" + cuerpo + "<br>" +
        "<a href='" + CONFIG_V2.SPREADSHEET_URL + "' style='display:inline-block;padding:12px 24px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:4px;font-weight:bold;'>Ver Spreadsheet</a></div>",
        name: CONFIG_V2.NOMBRE_LOCAL });
  } catch(err) { Logger.log("Error email accion: " + err.message); }
}

function instalarTriggerFacturas_V2() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onFacturaPendienteEdit_V2")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("onFacturaPendienteEdit_V2")
    .forSpreadsheet(CONFIG_V2.SPREADSHEET_ID).onEdit().create();
  Logger.log("Trigger onFacturaPendienteEdit_V2 instalado.");
}

function instalarTriggerSemanal_V2() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "triggerSemanalCC_V2")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("triggerSemanalCC_V2")
    .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(9).create();
  Logger.log("Trigger semanal V2 instalado.");
}

function instalarTriggerDiario_V2() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "actualizarEstadosCC_V2")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("actualizarEstadosCC_V2")
    .timeBased().everyDays(1).atHour(1).create();
  Logger.log("Trigger diario V2 instalado.");
}

function _aplicarFormatoVencido_V2(sheet, row, numCols) {
  const r = sheet.getRange(row, 1, 1, numCols);
  r.setBackground("#fce8e6"); r.setFontColor("#c0392b"); r.setFontWeight("bold");
}

function _limpiarFormatoFila_V2(sheet, row, numCols) {
  const r = sheet.getRange(row, 1, 1, numCols);
  r.setBackground(null); r.setFontColor(null); r.setFontWeight("normal");
}

function _actualizarEstadoPorId_V2(sheet, colId, id, colEstado, nuevoEstado) {
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][colId-1] || "").toString() === id) {
      sheet.getRange(i+1, colEstado).setValue(nuevoEstado);
      return true;
    }
  }
  return false;
}

function _actualizarEstadoPorIdMultiple_V2(sheet, colId, id, colEstado, nuevoEstado) {
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][colId-1] || "").toString() === id)
      sheet.getRange(i+1, colEstado).setValue(nuevoEstado);
  }
}

function _generarId_V2(prefijo) {
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const rand = Math.random().toString(36).substring(2, 6);
  return prefijo + "-" + fecha + "-" + rand;
}

function _armarLinkWhatsApp_V2(tel, mensaje) {
  let t = tel.toString().replace(/\D/g, "");
  if (t.length <= 10) t = "549" + t;
  return "https://wa.me/" + t + "?text=" + encodeURIComponent(mensaje);
}

function doGet(e) {
  try {
    const accion = e.parameter.accion;
    switch (accion) {
      case "buscarVentas": return _respGet(buscarVentas_V2(e.parameter.q || ""));
      case "getVenta":     return _respGet(getVenta_V2(e.parameter.id));
      case "ping":         return _respGet({ ok: true, msg: "MVP2 activo" });
      default:
        return HtmlService.createHtmlOutputFromFile("Index_MVP2")
          .setTitle("Hamilton Deco - Sistema POS")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    return _respGet({ ok: false, error: err.message });
  }
}

function _respGet(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function procesarDesdeHTML_V2(jsonStr) {
  const data = JSON.parse(jsonStr);
  const accion = data.accion;
  switch (accion) {
    case "registrarVenta":  return JSON.stringify(registrarVenta_V2(data));
    case "anularVenta":     return JSON.stringify(anularVenta_V2(data));
    case "registrarPagoCC": return JSON.stringify(registrarPagoCC_V2(data));
    case "getClientesCC":   return JSON.stringify(getClientesCC_V2());
    case "getListadoCC":    return JSON.stringify(getListadoCC_V2());
    case "buscarVentas":    return JSON.stringify(buscarVentas_V2(data.q || ""));
    case "getVenta":        return JSON.stringify(getVenta_V2(data.idVenta));
    case "getClientesBase": return JSON.stringify(getClientesBase_V2()); // ← NUEVO
    case "actualizarClienteBase": return JSON.stringify(actualizarClienteBase_V2(data)); // ← NUEVO
    default: return JSON.stringify({ ok: false, error: "Accion desconocida" });
  }
}

function onOpen() {
  _crearMenuDashboard();
}

// ← FIX v2.15: corrección de fechas históricas mal guardadas como Date object
function corregirFechasHistoricas_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const sh = ss.getSheetByName('VENTAS');
  const datos = sh.getDataRange().getValues();
  let corregidas = 0;

  for (let i = 1; i < datos.length; i++) {
    const fecha = datos[i][2];
    if (fecha instanceof Date) {
      const ts = datos[i][4];
      let fechaCorrecta = null;
      if (ts) {
        const tsStr = ts.toString();
        const m = tsStr.match(/(\d{4})-(\d{2})-(\d{2})/);
        if (m) {
          fechaCorrecta = m[3] + '/' + m[2] + '/' + m[1];
        }
      }
      if (fechaCorrecta) {
        _forzarFechaTexto_V2(sh, i + 1, 3, fechaCorrecta);
        Logger.log('Fila ' + (i+1) + ': corregida a ' + fechaCorrecta);
        corregidas++;
      }
    }
  }
  Logger.log('✅ Total corregidas: ' + corregidas);
}

// ← FIX v2.16: corrige fechas históricas mal guardadas como Date object
// en ITEMS, PAGOS, CC, PAGOS_CC y FACTURAS_PENDIENTES,
// usando VENTAS (ya corregida) como fuente de verdad por ID_VENTA.
// Ejecutar UNA SOLA VEZ desde el editor.
function corregirFechasHistoricas_Otras_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const shVentas = ss.getSheetByName('VENTAS');
  const datosVentas = shVentas.getDataRange().getValues();

  // Mapa ID_VENTA -> fecha correcta (string) desde VENTAS
  const fechaPorIdVenta = {};
  for (let i = 1; i < datosVentas.length; i++) {
    const id = (datosVentas[i][0] || '').toString().trim();
    const fecha = datosVentas[i][2];
    if (id && fecha) fechaPorIdVenta[id] = fecha.toString();
  }

  // Config: hoja, columna ID_VENTA (1-indexed), columna FECHA (1-indexed)
  const hojas = [
    { nombre: 'ITEMS',               colId: 1, colFecha: 2 },
    { nombre: 'PAGOS',               colId: 2, colFecha: 3 },
    { nombre: 'CC',                  colId: 2, colFecha: 3 },
    { nombre: 'FACTURAS_PENDIENTES', colId: 1, colFecha: 2 },
  ];

  hojas.forEach(function(h) {
    const sh = ss.getSheetByName(h.nombre);
    if (!sh || sh.getLastRow() < 2) return;
    const datos = sh.getDataRange().getValues();
    let corregidas = 0;

    for (let i = 1; i < datos.length; i++) {
      const fecha = datos[i][h.colFecha - 1];
      if (fecha instanceof Date) {
        const idVenta = (datos[i][h.colId - 1] || '').toString().trim();
        const fechaCorrecta = fechaPorIdVenta[idVenta];
        if (fechaCorrecta) {
          _forzarFechaTexto_V2(sh, i + 1, h.colFecha, fechaCorrecta);
          corregidas++;
        }
      }
    }
    Logger.log('✅ ' + h.nombre + ': ' + corregidas + ' fechas corregidas');
  });

  // PAGOS_CC no tiene ID_VENTA, solo NOMBRE_CLIENTE — se corrige aparte,
  // usando el helper genérico (reconstruye desde el propio valor si es Date,
  // asumiendo que Sheets no pudo invertir DD/MM si el día es >12;
  // para días ambiguos, se deja para revisión manual y se loggea).
  const shPagosCC = ss.getSheetByName('PAGOS_CC');
  if (shPagosCC && shPagosCC.getLastRow() >= 2) {
    const datosPagosCC = shPagosCC.getDataRange().getValues();
    let corregidasPCC = 0, revisar = 0;
    for (let i = 1; i < datosPagosCC.length; i++) {
      const fecha = datosPagosCC[i][2]; // columna C
      if (fecha instanceof Date) {
        const dd = fecha.getDate();
        if (dd > 12) {
          // Sheets no pudo haber invertido: dd es válido como día, mm como mes
          const fechaStr = Utilities.formatDate(fecha, CONFIG_V2.TZ, 'dd/MM/yyyy');
          _forzarFechaTexto_V2(shPagosCC, i + 1, 3, fechaStr);
          corregidasPCC++;
        } else {
          Logger.log('⚠️ PAGOS_CC fila ' + (i+1) + ': fecha ambigua, revisar manualmente: ' + fecha);
          revisar++;
        }
      }
    }
    Logger.log('✅ PAGOS_CC: ' + corregidasPCC + ' corregidas, ' + revisar + ' requieren revisión manual');
  }
}

// ← FIX v2.16b: corrige PAGOS_CC usando el propio ID_PAGO_CC como fuente de verdad
// (el ID se generó con la fecha real, nunca sufrió reinterpretación de Sheets)
function corregirFechasPagosCC_V2() {
  const ss = SpreadsheetApp.openById(CONFIG_V2.SPREADSHEET_ID);
  const sh = ss.getSheetByName('PAGOS_CC');
  if (!sh || sh.getLastRow() < 2) { Logger.log('❌ PAGOS_CC vacía o no encontrada'); return; }

  const datos = sh.getDataRange().getValues();
  let corregidas = 0;

  for (let i = 1; i < datos.length; i++) {
    const id = (datos[i][0] || '').toString().trim();
    const m = id.match(/PCC-(\d{4})(\d{2})(\d{2})-/);
    if (!m) continue;
    const fechaCorrecta = m[3] + '/' + m[2] + '/' + m[1]; // DD/MM/YYYY
    _forzarFechaTexto_V2(sh, i + 1, 3, fechaCorrecta);
    corregidas++;
  }
  Logger.log('✅ PAGOS_CC: ' + corregidas + ' filas normalizadas desde su ID (fuente de verdad)');
}
