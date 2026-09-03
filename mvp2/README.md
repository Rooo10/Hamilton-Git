# Hamilton MVP2 — Backup de código

Proyecto GAS activo: `Hamilton_MVP2` — cuenta `HamiltonDecoRosario@gmail.com`
Spreadsheet ID: `18RlCDXP0XEXwtQvWhOO4yHLQWEfdtrmRTWgs698qIv8`
Deploy URL: https://script.google.com/macros/s/AKfycbwW0OK0cOhyGEVFVikeKCA68gfWer07VtjPDcFuFtv-EQ6qvkTB8PgvGA85aQhm3j1g0Q/exec

## Estado de este backup (02/09/2026)

## Estado de este backup (03/09/2026)

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

Los 5 archivos activos de MVP2 fueron confirmados 1:1 contra el código real del
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
