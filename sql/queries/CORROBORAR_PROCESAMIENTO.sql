-- ==========================================================
-- Consulta : CORROBORAR_PROCESAMIENTO
-- Exportado: 2026-10-22 22:23:13
-- ==========================================================

SELECT DBIngredientes.Nombre, Cotizaciones.Unidad, Procesamiento.Cantidad, Procesamiento.MedidaInicial, Procesamiento.MedidaFinal, Procesamiento.Fecha, Procesamiento.Observaciones, Cotizaciones.CodCotizacion
FROM DBIngredientes INNER JOIN (Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente;

