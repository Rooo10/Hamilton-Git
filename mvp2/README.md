# Hamilton MVP2 — Backup de código

Proyecto GAS activo: `Hamilton_MVP2` — cuenta `HamiltonDecoRosario@gmail.com`
Spreadsheet ID: `18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8`
Deploy URL: https://script.google.com/macros/s/AKfycbwW0OK0cOhyGEVFVikeKCA68gfWer07VtjPDcFuFtv-EQ6qvkTB8PgvGA85aQhm3j1g0Q/exec

## Estado de este backup (18/09/2026)

| Archivo | Versión funcional | Última actualización |
|---|---|---|
| `Codigo_MVP2_V2.gs` | v2.23 — rutas `cierreDeCaja` y `getUltimoCierre` en `procesarDesdeHTML_V2`; se quitó la `cierreDeCajaHTML_V2` vieja (ahora vive en `CierrePanel_MVP2.gs`) | 18/09/2026 |
| `CierreDiario_MVP2.gs` | v2.22 — `completarVentasDiarias_VD` ubica las tablas de `VENTAS_DIARIAS` por su encabezado "Mes" (`_ubicarTablas_VD`), sin números de fila fijos | 18/09/2026 |
| `CierrePanel_MVP2.gs` | **Nuevo** (v2.22) — backend del panel "Cierre del día": `cierreDeCajaHTML_V2`, `getUltimoCierreHTML_V2` y la foto guardada del cierre | 18/09/2026 |
| `Index_MVP2.html` | v2.23 — botón y panel "Cierre del día" + ventanas propias en lugar de `alert`/`confirm` | 18/09/2026 |
| `Dashboard_MVP2.gs` | Sin cambios desde v2.21 | 03/09/2026 |
| `InitSheet_MVP2_V3.gs` | Sin cambios desde el 31/08 | 03/09/2026 |
| `BaseClientes_MVP2.gs` | Sin cambios desde julio | 03/09/2026 |

Los cambios del 18/09 se aplicaron a este backup con las mismas ediciones hechas en el
editor de Apps Script, y se probaron en `/dev` y en `/exec`.

### Cierre del día — cómo funciona
- Botón "🧾 Cierre del día" en la pantalla de caja. Al abrir el panel se muestra primero el
  último cierre guardado y enseguida se recalcula solo (`calcularCierreDiario()` +
  `completarVentasDiarias_VD()`), así la cajera puede corregir algo y volver a correrlo.
- El panel muestra: indicadores del día (`CIERRE_DIARIO`, filas 4 a 17), ítems vendidos y
  pagos por forma de pago del día con su total (`VENTAS_DIARIAS`, Tabla Diaria).
- Cada cálculo guarda una "foto" (indicadores, ítems, pagos y fecha/hora del cálculo) en
  Script Properties, en trozos de 3000 caracteres (`CIERRE_FOTO_N`, `CIERRE_FOTO_0…`).
- Las ventas anuladas no cuentan en ningún bloque.

### Changelog de esta sesión (18/09/2026)
- v2.22 → Cierre del día: panel en `Index_MVP2.html`, backend nuevo `CierrePanel_MVP2.gs`,
  rutas `cierreDeCaja` y `getUltimoCierre`.
- v2.22 → `VENTAS_DIARIAS`: se eliminó la leyenda (filas 2 a 13) y `completarVentasDiarias_VD`
  dejó de depender de filas fijas.
- v2.23 → Ventanas propias (`mostrarAviso`, `mostrarConfirmacion`) en lugar de los carteles
  del navegador: 21 `alert()` y 2 `confirm()` reemplazados.
- v2.23 → Fix: el cartel "Venta registrada" mostraba "Efectivo $0" si se elegía la forma de
  pago antes de cargar los productos (el importe guardado siempre fue correcto).
- (16/09) Fix de días de mora (`_parsearFechaDDMMYYYY_V2`) y ruta `cierreDeCaja`.

⚠️ `normalizarFechasVentasDiarias_VD` y `agregarTotalVentasDiarias_VD` (migraciones de una
sola vez, ya ejecutadas) todavía usan filas fijas de la disposición anterior de
`VENTAS_DIARIAS`. No volver a correrlas.

## Estado anterior del backup (03/09/2026)

| Archivo | Versión funcional | Última confirmación |
|---|---|---|
| `Codigo_MVP2_V2.gs` | v2.21 (getClientesBase/actualizarClienteBase agregados al switch) | 03/09/2026 |
| `Dashboard_MVP2.gs` | v2.21 (incluye getClientesBase_V2 y actualizarClienteBase_V2 al final del archivo — comparten espacio global con Codigo_MVP2_V2.gs sin conflicto) | 03/09/2026 |
| `Index_MVP2.html` | v2.21 (buscador/autocompletar de clientes en el formulario de venta + botón y panel "Datos Clientes" con edición) | 03/09/2026 |
| `InitSheet_MVP2_V3.gs` | Sin cambios desde el 31/08 | 03/09/2026 (re-verificado, intacto) |
| `BaseClientes_MVP2.gs` | Sin cambios desde julio (`actualizarBaseYAsignarIDs_D`, `triggerBaseClientes`) | 03/09/2026 (re-verificado, intacto) |

Nota: `LISTADO_CC.gs` no existe como archivo separado en el proyecto GAS —
confirmado visualmente en el editor (03/09/2026). Era un ítem fantasma del plan
original de mayo; la lógica de LISTADO_CC vive dentro de `Codigo_MVP2_V2.gs`
(`_reconstruirListadoCC_V2`, `getListadoCC_V2`). Retirado del checklist.

Al 03/09/2026, los 5 archivos activos de MVP2 fueron confirmados 1:1 contra el código real del
editor de Apps Script el 03/09/2026, pasados dos veces por el usuario para
descartar mezclas entre pestañas. Backup completo.

### Changelog de esta sesión (02/09/2026)
- v2.17 → Tarjeta Naranja como método de pago #15 (crédito, con cuotas)
- (fix) `doGet` recuperado — se había perdido al limpiar archivos obsoletos del proyecto
- v2.18 → Fecha de venta editable, obligatoria, con bloqueo si queda vacía
- v2.19 → Foco inicial en Nombre del Cliente (antes saltaba a Descripción de producto)
- v2.20 → Opción "No enviar" en el desplegable de envío de factura + detalle de ítems
  (cantidad, descripción, subtotal) en el mensaje de WhatsApp del recibo

## Archivos obsoletos detectados en el proyecto GAS (no incluidos aquí)

El editor de Apps Script (`Hamilton_MVP2`) tiene varios archivos legacy que ya no están
en uso y generan confusión en el menú:
- `Código.gs` — vacío, reservado
- `Codigo_MVP2.gs` — versión pre-V2, reemplazada por `Codigo_MVP2_V2.gs`
- `InitSheet_MVP2.gs` — reemplazada por V3
- `InitSheet_MVP2_V2.gs` — reemplazada por V3

Ver recomendación de limpieza antes de borrarlos.
