-- ==========================================================
-- Consulta : ControlIngresosCompras
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT Cotizaciones.CodCotizacion, DBIngredientes.Nombre, Cotizaciones.Marca, Cotizaciones.Capacidad, Compras.Cantidad, numerosemana([Compras]![Fecha]) AS semana
FROM DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente
WHERE (((numerosemana(Compras!Fecha))=[semana]) And ((DBIngredientes.TIPO1)="VARIABLES"));

