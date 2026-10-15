-- ==========================================================
-- Consulta : MesSistema
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

SELECT Month([Dates]) AS mes, Year([Dates]) AS año
FROM FechaSistema
GROUP BY Month([Dates]), Year([Dates]);

