-- ==========================================================
-- Consulta : HistorialComprasDatosCentral
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT Compras.Fecha, Cotizaciones.CodIngrediente, Compras.CodCotizacion, Compras.NumeroFactura, nombreproductocotiprocesado([Compras]![CodCotizacion]) AS NombreProd, Cotizaciones.Conversion, Cotizaciones.ConversionEstandar, DBIngredientes.NombreSinProcesar, Compras.Pagado, Compras.Cantidad, Compras.CostoTotal, Compras.Observaciones, DBIngredientes.Tipo, DBIngredientes.TIPO1, DBIngredientes.TIPO2, Compras.Tipo, codigolocal() AS [local], Proovedores.TipoDePago, Compras.Peso, Compras.Lote, Proovedores.Nombre, [Compras]![Destino] AS CeCo
FROM (Compras INNER JOIN (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON Compras.CodCotizacion = Cotizaciones.CodCotizacion) INNER JOIN Proovedores ON Compras.CodProveedor = Proovedores.CodProovedor
WHERE (((Compras.Fecha) Between [Formularios]![Menu Modulo Contabilidad]![excelcomprasdesde] And [Formularios]![Menu Modulo Contabilidad]![excelcomprashasta]))
ORDER BY Compras.Fecha;

