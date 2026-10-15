-- ==========================================================
-- Consulta : PanYPaz_Ingresos
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

SELECT IngresosPitaya.CodCotizacion, IngresosPitaya.Fecha, DBIngredientes.Nombre, IngresosPitaya.Cantidad, DBIngredientes.CodIngrediente
FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion
WHERE (((DBIngredientes.Tipo)="Pan y Paz"))
ORDER BY IngresosPitaya.Fecha, IngresosPitaya.Cantidad, DBIngredientes.CodIngrediente;

