-- ==========================================================
-- Consulta : HistorialInventariosDatos
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT numerosemana([Inventario Cotizacion]![Fecha]) AS aSemana, Cotizaciones.Conversion, Cotizaciones.Unidad, DBIngredientes.NombreSinProcesar, nombreproductocotiprocesado([Inventario Cotizacion]![CodCotizacion]) AS Nombre, [Inventario Cotizacion].Fecha, Sum([Inventario Cotizacion].Cantidad) AS SumaDeCantidad, DBIngredientes.Tipo, DBIngredientes.TIPO1, DBIngredientes.TIPO2, [Inventario Cotizacion].CodCotizacion, codigolocal() AS [local], PorcionDentroDeMezcla([Inventario Cotizacion]![CodCotizacion]) AS mezcla
FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN [Inventario Cotizacion] ON Cotizaciones.CodCotizacion = [Inventario Cotizacion].CodCotizacion
GROUP BY numerosemana([Inventario Cotizacion]![Fecha]), Cotizaciones.Conversion, Cotizaciones.Unidad, DBIngredientes.NombreSinProcesar, nombreproductocotiprocesado([Inventario Cotizacion]![CodCotizacion]), [Inventario Cotizacion].Fecha, DBIngredientes.Tipo, DBIngredientes.TIPO1, DBIngredientes.TIPO2, [Inventario Cotizacion].CodCotizacion, codigolocal(), PorcionDentroDeMezcla([Inventario Cotizacion]![CodCotizacion])
HAVING (((numerosemana([Inventario Cotizacion]![Fecha]))=[Formularios]![Menu Gestion]![invseman]) AND ((PorcionDentroDeMezcla([Inventario Cotizacion]![CodCotizacion]))=0))
ORDER BY [Inventario Cotizacion].Fecha;

