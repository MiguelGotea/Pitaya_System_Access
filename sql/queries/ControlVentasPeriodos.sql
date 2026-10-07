-- ==========================================================
-- Consulta : ControlVentasPeriodos
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT NotaDePedido.Fecha, Hour([NotaDePedido]![Hora]) AS Hora, Sum(FactorTiempoProduccion([SubPedido]![CodSubPedido])) AS Monto
FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido
GROUP BY NotaDePedido.Fecha, Hour([NotaDePedido]![Hora])
HAVING (((NotaDePedido.Fecha)>=[Formularios]![ControlVentasPeriodos]![fechainicio] And (NotaDePedido.Fecha)<=#7/12/2018#) AND ((Hour([NotaDePedido]![Hora]))>=7 And (Hour([NotaDePedido]![Hora]))<=21));

