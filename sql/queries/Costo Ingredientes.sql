-- ==========================================================
-- Consulta : Costo Ingredientes
-- Exportado: 2026-10-07 06:21:32
-- ==========================================================
SELECT DBIngredientes.Nombre AS Expr1, DBIngredientes.CodIngrediente AS Expr2, CostoIngredienteGlobal([DBIngredientes]![CodIngrediente],74) AS Costo, DBIngredientes.TIPO1 AS Expr3, DBIngredientes.Tipo AS Expr4
FROM DBIngredientes
WHERE (((DBIngredientes.TIPO1)="VARIABLES") And ((DBIngredientes.Tipo)="Frutas"));

