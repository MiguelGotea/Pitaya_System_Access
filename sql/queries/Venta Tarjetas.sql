-- ==========================================================
-- Consulta : Venta Tarjetas
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT NotaDePedido.Fecha, NotaDePedido.Hora, SubPedido.CodBatido, SubPedido.Cantidad, SubPedido.Observaciones, NotaDePedido.CodCliente, NotaDePedido.CodPedido
FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido
WHERE (((SubPedido.CodBatido)="tcp"))
ORDER BY NotaDePedido.Fecha DESC;

