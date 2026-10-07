-- ==========================================================
-- Consulta : ParetoIngredientes
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT Sum(Compras.CostoTotal) AS SumaDeCostoTotal, DBIngredientes.Nombre, DBIngredientes.Unidad
FROM DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente
GROUP BY DBIngredientes.Nombre, DBIngredientes.Unidad, numerosemana([Compras]![Fecha])<numerosemana(Date()), numerosemana([Compras]![Fecha])>numerosemana(Date())-7, DBIngredientes.TIPO1
HAVING (((numerosemana(Compras!Fecha)<numerosemana(Date()))=True) And ((numerosemana(Compras!Fecha)>numerosemana(Date())-7)=True) And ((DBIngredientes.TIPO1)="VARIABLES"))
ORDER BY Sum(Compras.CostoTotal) DESC;

