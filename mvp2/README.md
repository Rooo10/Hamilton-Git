# Hamilton MVP2 — Backup de código

Proyecto GAS activo: `Hamilton_MVP2` — cuenta `HamiltonDecoRosario@gmail.com`
Spreadsheet ID: `18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8`
Deploy URL: https://script.google.com/macros/s/AKfycbwW0OK0cOhyGEVFVikeKCA68gfWer07VtjPDcFuFtv-EQ6qvkTB8PgvGA85aQhm3j1g0Q/exec

## Estado de este backup (02/09/2026)

| Archivo | Versión funcional | Última confirmación |
|---|---|---|
| `Codigo_MVP2_V2.gs` | v2.18 (Tarjeta Naranja + doGet recuperado + fecha de venta editable) | 02/09/2026 |
| `Dashboard_MVP2.gs` | v2.17 (15 métodos de pago, columnas Q/R/S) | 31/08/2026 |
| `Index_MVP2.html` | v2.20 (Tarjeta Naranja, fecha editable obligatoria, foco inicial en Nombre, "No enviar" en factura, detalle de ítems en WhatsApp) | 02/09/2026 |
| `InitSheet_MVP2_V3.gs` | Sin cambios desde el 31/08 | 31/08/2026 |
| `BaseClientes_MVP2.gs` | Sin cambios desde el 31/08 | 31/08/2026 |
| `LISTADO_CC.gs` | ⏳ Falta | No se pudo confirmar contenido actual |

Los 5 archivos activos de MVP2 quedaron confirmados 1:1 contra el código real en el
editor de Apps Script. Solo falta `LISTADO_CC.gs` para tener el backup completo.

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
