-- ==========================================================
-- Consulta : Detalle por Pedido
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT NotaDePedido.CodPedido AS Expr1, NotaDePedido.Fecha AS Expr2, NotaDePedido.Hora AS Expr3, UsoVidrio([NotaDePedido]![CodPedido]) AS Envase, NotaDePedido.CodCliente AS CLUB, NotaDePedido.Edad AS Expr4, NotaDePedido.Nacionalidad AS Expr5, NotaDePedido.Personas AS Expr6, NotaDePedido.CantidadClientes AS Expr7, NotaDePedido.PrimeraVez AS Expr8, MontoPedido([NotaDePedido]![CodPedido]) AS Monto
FROM NotaDePedido
WHERE ((([NotaDePedido].[Fecha])>=#10/19/2016#));

