-- ==========================================================
-- Consulta : PanYPaz_Venta
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT SubPedido.CodSubPedido, SubPedido.CodBatido, NotaDePedido.Fecha, DBBatidos.Nombre, SubPedido.Cantidad, DBBatidos.Precio, SubReceta.CodIngrediente, SubPedido.CodPromocion
FROM ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido
WHERE (((DBBatidos.CodGrupo)=10))
ORDER BY NotaDePedido.Fecha, SubPedido.Cantidad, SubReceta.CodIngrediente;

