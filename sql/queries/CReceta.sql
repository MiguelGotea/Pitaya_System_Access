-- ==========================================================
-- Consulta : CReceta
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

SELECT DBBatidos.Nombre, DBBatidos.Medida, DBIngredientes.Nombre, SubReceta.Cantidad, DBIngredientes.Unidad, DBBatidos.CodBatido
FROM DBIngredientes INNER JOIN (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente;

