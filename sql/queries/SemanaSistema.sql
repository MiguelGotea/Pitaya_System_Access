-- ==========================================================
-- Consulta : SemanaSistema
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

SELECT numerosemana([Dates]) AS semana
FROM FechaSistema
GROUP BY numerosemana([Dates])
ORDER BY numerosemana([Dates]);

