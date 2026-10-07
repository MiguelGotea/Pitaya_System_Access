-- ==========================================================
-- Consulta : InformeDiario
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT SubPedido.CodPedido, NotaDePedido.Fecha, UsoVidrio([NotaDePedido]![CodPedido]) AS Envase, NotaDePedido.Hora, DBBatidos.Nombre, DBBatidos.Medida, DBBatidos.Precio, SubPedido.Cantidad, PrecioReal([SubPedido]![CodSubPedido])*[SubPedido]![Cantidad] AS Sub, SubPedido.Observaciones, DBPromociones.Nombre, NotaDePedido.CodCliente, NotaDePedido.Modalidad
FROM DBPromociones INNER JOIN (NotaDePedido INNER JOIN (DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido) ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBPromociones.CodPromocion = SubPedido.CodPromocion
WHERE (((NotaDePedido.Fecha)=Date()));

