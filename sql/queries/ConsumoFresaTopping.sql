-- ==========================================================
-- Consulta : ConsumoFresaTopping
-- Exportado: 2026-10-07 06:21:32
-- ==========================================================
SELECT SubReceta.CodIngrediente, SubReceta.Tipo, numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()) And numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-5 AS rango, NotaDePedido.Anulado, Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]) AS TotalGramos, "4 semanas" AS Tota
FROM (SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido) INNER JOIN (SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) ON SubPedido.CodBatido = DBBatidos.CodBatido
GROUP BY SubReceta.CodIngrediente, SubReceta.Tipo, numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()) And numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-5, NotaDePedido.Anulado, "4 semanas"
HAVING (((SubReceta.CodIngrediente)="F028") AND ((SubReceta.Tipo)="T") AND ((numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()) And numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-5)=-1) AND ((NotaDePedido.Anulado)=False));

