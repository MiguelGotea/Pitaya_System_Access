-- ==========================================================
-- Consulta : ListaPORCIONESFiltro
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion]) AS nombrefinal, Grupos.control, TiposVariables.Control, DBBatidos.Vigencia, SubReceta.codporcion, [DBBatidos]![CodGrupo]=7 AS nomostrador, PorcionDentroDeMezcla([SubReceta]![codporcion]) AS mezcla, DBIngredientes.Tipo, DBIngredientes.Nombre
FROM (TiposVariables INNER JOIN ((SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente) ON TiposVariables.Tipo = DBIngredientes.Tipo) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo
GROUP BY TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion]), Grupos.control, TiposVariables.Control, DBBatidos.Vigencia, SubReceta.codporcion, [DBBatidos]![CodGrupo]=7, PorcionDentroDeMezcla([SubReceta]![codporcion]), DBIngredientes.Tipo, DBIngredientes.Nombre
HAVING (((Grupos.control)=True) AND ((TiposVariables.Control)=True) AND ((DBBatidos.Vigencia)=True) AND ((SubReceta.codporcion) Is Not Null) AND (([DBBatidos]![CodGrupo]=7)=0) AND ((PorcionDentroDeMezcla([SubReceta]![codporcion]))=0))
ORDER BY TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion]);

