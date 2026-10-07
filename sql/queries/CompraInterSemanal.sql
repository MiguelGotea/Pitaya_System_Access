-- ==========================================================
-- Consulta : CompraInterSemanal
-- Exportado: 2026-10-07 06:21:32
-- ==========================================================
SELECT DBIngredientes.Nombre, Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*FactorDeUso([SubReceta]![CodIngrediente],[SubPedido]![CodPedido])*(StockDeSeguridad([DBIngredientes]![CodIngrediente])+1)/3) AS PEDIDO, DBIngredientes.Unidad, Cotizaciones.Marca, Cotizaciones.Capacidad, Cotizaciones.Unidad, DBIngredientes.CodIngrediente
FROM ((DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBBatidos.CodBatido = SubPedido.CodBatido) INNER JOIN (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) ON DBBatidos.CodBatido = SubReceta.CodBatido) INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente
GROUP BY DBIngredientes.Nombre, DBIngredientes.Unidad, Cotizaciones.Marca, Cotizaciones.Capacidad, Cotizaciones.Unidad, DBIngredientes.CodIngrediente, DBIngredientes.Tipo, Cotizaciones.Prioridad, numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()), numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-4
HAVING (((Cotizaciones.Prioridad)=True) And ((numerosemana(NotaDePedido!Fecha)<numerosemana(Date()))=-1) And ((numerosemana(NotaDePedido!Fecha)>numerosemana(Date())-4)=-1))
ORDER BY DBIngredientes.Tipo;

