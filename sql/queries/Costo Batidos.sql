-- ==========================================================
-- Consulta : Costo Batidos
-- Exportado: 2026-10-07 06:21:32
-- ==========================================================
SELECT DBBatidos.CodBatido AS Expr1, DBBatidos.Nombre AS Expr2, DBBatidos.Medida AS Expr3, DBBatidos.Precio AS Expr4, CostoTotalBatido([DBBatidos]![CodBatido],[semana]) AS Costo, DBBatidos.Vigencia AS Expr5
FROM DBBatidos
WHERE ((([DBBatidos].[Vigencia])=True));

