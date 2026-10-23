-- ==========================================================
-- Consulta : PanYPaz_Compra
-- Exportado: 2026-10-22 22:23:14
-- ==========================================================

SELECT Compras.CodCotizacion, Compras.Fecha, DBIngredientes.Nombre, Compras.Cantidad, Cotizaciones.CodIngrediente, Compras.CostoTotal
FROM DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente
WHERE (((DBIngredientes.Tipo)="Pan y Paz"))
ORDER BY Compras.Fecha, Compras.Cantidad, Cotizaciones.CodIngrediente;

