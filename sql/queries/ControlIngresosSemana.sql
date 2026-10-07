-- ==========================================================
-- Consulta : ControlIngresosSemana
-- Exportado: 2026-10-07 06:21:32
-- ==========================================================
SELECT Cotizaciones.CodCotizacion, DBIngredientes.Nombre, Cotizaciones.Marca, Cotizaciones.Capacidad, DBIngredientes.TIPO1, IngresosPitaya.Cantidad, numerosemana([IngresosPitaya]![Fecha]) AS semana
FROM IngresosPitaya INNER JOIN (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion
WHERE (((DBIngredientes.TIPO1)="VARIABLES") And ((numerosemana(IngresosPitaya!Fecha))=[semana]));

