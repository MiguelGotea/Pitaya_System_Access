-- ==========================================================
-- Consulta : MesSistema
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT Month([Dates]) AS mes, Year([Dates]) AS año
FROM FechaSistema
GROUP BY Month([Dates]), Year([Dates]);

