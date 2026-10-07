-- ==========================================================
-- Consulta : Egresos Costos Fijos
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT Compras.Fecha, Month([Compras]![Fecha]) AS Mes, numerosemana([Compras]![Fecha]) AS Semana, DBIngredientes.TIPO1, DBIngredientes.TIPO2, DBIngredientes.Tipo AS TIPO3, [DBIngredientes]![Nombre] & " " & [Compras]![Observaciones] AS Concepto, Cotizaciones.Marca, Proovedores.Nombre, Compras.Cantidad, [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad] AS PRESENTACION, ConversionCalculado([Compras]![CodCotizacion],numerosemana([Compras]![Fecha]))*[Compras]![Cantidad] AS [CANTIDAD CONTROL], DBIngredientes.Unidad, Compras.CostoTotal, [Compras]![CostoTotal]/[Compras]![Cantidad] AS CUU, [Compras]![CostoTotal]/(ConversionCalculado([Compras]![CodCotizacion],numerosemana([Compras]![Fecha]))*[Compras]![Cantidad]) AS CUF, Cotizaciones.CodIngrediente, Compras.Destino
FROM Proovedores INNER JOIN (DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON Proovedores.CodProovedor = Compras.CodProveedor
WHERE (((Month(Compras!Fecha))=3) And ((DBIngredientes.TIPO1)="FIJOS"))
ORDER BY Compras.Fecha;

