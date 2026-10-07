-- ==========================================================
-- Consulta : Consumo semanal
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT DBIngredientes.Nombre, Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]) AS Cantidad, DBIngredientes.Unidad, numerosemana([NotaDePedido]![Fecha]) AS Semana, DBIngredientes.Tipo
FROM DBIngredientes INNER JOIN ((DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBBatidos.CodBatido = SubPedido.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente
GROUP BY DBIngredientes.Nombre, DBIngredientes.Unidad, numerosemana([NotaDePedido]![Fecha]), DBIngredientes.Tipo;

