-- ==========================================================
-- Consulta : VentaUnidadesDia
-- Exportado: 2026-10-07 06:21:34
-- ==========================================================
SELECT numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()) AS Semana, Hour([NotaDePedido]![Hora]) AS Hora, Weekday([NotaDePedido]![Fecha],2) AS Dia, numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-6 AS Semana2, Sum(FactorTiempoProduccion([SubPedido]![CodSubPedido])) AS Produccion
FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido
GROUP BY numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()), Hour([NotaDePedido]![Hora]), Weekday([NotaDePedido]![Fecha],2), numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-6
HAVING (((numerosemana(NotaDePedido!Fecha)<numerosemana(Date()))=-1) And ((numerosemana(NotaDePedido!Fecha)>numerosemana(Date())-6)=-1));

