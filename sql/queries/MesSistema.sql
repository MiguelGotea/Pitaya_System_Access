-- ==========================================================
-- Consulta : MesSistema
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT Month([Dates]) AS mes, Year([Dates]) AS año
FROM FechaSistema
GROUP BY Month([Dates]), Year([Dates]);

