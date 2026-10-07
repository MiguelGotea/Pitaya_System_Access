-- ==========================================================
-- Consulta : InventarioEIngresos
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT CodICotizacion, CodCotizacion, Cantidad, Fecha+1 as FechaS, "E" as Tipo FROM [Inventario Cotizacion]
UNION SELECT CodIngresoPitaya, CodCotizacion, Cantidad, Fecha as FechaS, "I" as Tipo FROM IngresosPitaya;

