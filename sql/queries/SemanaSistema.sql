-- ==========================================================
-- Consulta : SemanaSistema
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT numerosemana([Dates]) AS semana
FROM FechaSistema
GROUP BY numerosemana([Dates])
ORDER BY numerosemana([Dates]);

