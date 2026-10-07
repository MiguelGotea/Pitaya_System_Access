-- ==========================================================
-- Consulta : Venta Batidos
-- Exportado: 2026-10-07 06:21:34
-- ==========================================================
SELECT DBBatidos.Nombre, Sum(SubPedido.Cantidad) AS SumaDeCantidad, numerosemana([NotaDePedido]![Fecha]) AS Semana, Grupos.NombreGrupo
FROM Grupos INNER JOIN (DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBBatidos.CodBatido = SubPedido.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo
GROUP BY DBBatidos.Nombre, numerosemana([NotaDePedido]![Fecha]), Grupos.NombreGrupo;

