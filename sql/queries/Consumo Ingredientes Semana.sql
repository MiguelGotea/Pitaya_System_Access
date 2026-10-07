-- ==========================================================
-- Consulta : Consumo Ingredientes Semana
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT DBIngredientes.Nombre, Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]/4) AS Consumo
FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN (DBBatidos INNER JOIN (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) ON DBBatidos.CodBatido = SubReceta.CodBatido) ON SubPedido.CodBatido = DBBatidos.CodBatido
GROUP BY DBIngredientes.Nombre, numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()), numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-5, DBIngredientes.TIPO1
HAVING (((numerosemana(NotaDePedido!Fecha)<numerosemana(Date()))=True) And ((numerosemana(NotaDePedido!Fecha)>numerosemana(Date())-5)=True) And ((DBIngredientes.TIPO1)="VARIABLES"))
ORDER BY Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]/4) DESC;

