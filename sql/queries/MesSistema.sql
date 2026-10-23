-- ==========================================================
-- Consulta : MesSistema
-- Exportado: 2026-10-22 22:23:14
-- ==========================================================

SELECT Month([Dates]) AS mes, Year([Dates]) AS año
FROM FechaSistema
GROUP BY Month([Dates]), Year([Dates]);

