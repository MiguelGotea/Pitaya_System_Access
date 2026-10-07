-- ==========================================================
-- Consulta : SemanaSistema
-- Exportado: 2026-10-07 06:21:34
-- ==========================================================
SELECT numerosemana([Dates]) AS semana
FROM FechaSistema
GROUP BY numerosemana([Dates])
ORDER BY numerosemana([Dates]);

