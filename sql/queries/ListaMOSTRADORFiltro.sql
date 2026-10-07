-- ==========================================================
-- Consulta : ListaMOSTRADORFiltro
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT DBBatidos.Marca, DBBatidos.CodSubGrupo, DBBatidos.CodGrupo, DBBatidos.Vigencia, SubReceta.InsumoClave, IIf(IsNull([SubReceta]![codporcion]),DLookUp("[Nombre]","[DBIngredientes]","[CodIngrediente]='" & [SubReceta]![CodIngrediente] & "'"),nombreproductocotiprocesado([SubReceta]![codporcion])) AS nombreprod, SubReceta.codporcion, SubReceta.CodIngrediente, CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido]) AS coti
FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido
GROUP BY DBBatidos.Marca, DBBatidos.CodSubGrupo, DBBatidos.CodGrupo, DBBatidos.Vigencia, SubReceta.InsumoClave, IIf(IsNull([SubReceta]![codporcion]),DLookUp("[Nombre]","[DBIngredientes]","[CodIngrediente]='" & [SubReceta]![CodIngrediente] & "'"),nombreproductocotiprocesado([SubReceta]![codporcion])), SubReceta.codporcion, SubReceta.CodIngrediente, CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido])
HAVING (((DBBatidos.CodGrupo)=7) AND ((DBBatidos.Vigencia)=True) AND ((SubReceta.InsumoClave)=True) AND ((SubReceta.codporcion) Is Not Null))
ORDER BY DBBatidos.Marca, DBBatidos.CodSubGrupo;

