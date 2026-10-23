-- ==========================================================
-- Consulta : SemanaSistema
-- Exportado: 2026-10-22 22:23:14
-- ==========================================================

SELECT numerosemana([Dates]) AS semana
FROM FechaSistema
GROUP BY numerosemana([Dates])
ORDER BY numerosemana([Dates]);

