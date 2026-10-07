-- ==========================================================
-- Consulta : ProyeccionVentas
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT numerosemana([NotaDePedido]![Fecha]) AS semana, Sum(SubPedido.Cantidad) AS SumaDeCantidad, DBBatidos.Nombre, Grupos.NombreGrupo, numerosemana(Date())=numerosemana([NotaDePedido]![Fecha]) AS EstaSemana
FROM Grupos INNER JOIN ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo
GROUP BY numerosemana([NotaDePedido]![Fecha]), DBBatidos.Nombre, Grupos.NombreGrupo, numerosemana(Date())=numerosemana([NotaDePedido]![Fecha])
HAVING (((numerosemana(NotaDePedido!Fecha))>numerosemana(Date())-8) And ((Grupos.NombreGrupo)<>"PAN Y PAZ" And (Grupos.NombreGrupo)<>"ADICIONALES" And (Grupos.NombreGrupo)<>"TARJETA") And ((numerosemana(Date())=numerosemana(NotaDePedido!Fecha))=0))
ORDER BY Sum(SubPedido.Cantidad) DESC;

