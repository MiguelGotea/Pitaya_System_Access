-- ==========================================================
-- Consulta : CORROBORAR_PROCESAMIENTO
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT DBIngredientes.Nombre, Cotizaciones.Unidad, Procesamiento.Cantidad, Procesamiento.MedidaInicial, Procesamiento.MedidaFinal, Procesamiento.Fecha, Procesamiento.Observaciones, Cotizaciones.CodCotizacion
FROM DBIngredientes INNER JOIN (Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente;

