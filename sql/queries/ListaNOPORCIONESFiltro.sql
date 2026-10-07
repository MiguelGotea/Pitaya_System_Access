-- ==========================================================
-- Consulta : ListaNOPORCIONESFiltro
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control, [DBBatidos]![CodGrupo]=7 AS nomostrador, DBBatidos.Vigencia, Cotizaciones.Subproducto, SubReceta.codporcion, [Cotizaciones]![Marca] & " "="Almacen Global " AS NoAlmacenGlobal, Cotizaciones.Prioridad, Cotizaciones.Descontinuado, Cotizaciones.CodCotizacion
FROM ((((DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente) INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo) INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo
GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control, [DBBatidos]![CodGrupo]=7, DBBatidos.Vigencia, Cotizaciones.Subproducto, SubReceta.codporcion, [Cotizaciones]![Marca] & " "="Almacen Global ", Cotizaciones.Prioridad, Cotizaciones.Descontinuado, Cotizaciones.CodCotizacion
HAVING (((Grupos.control)=True) AND ((TiposVariables.Control)=True) AND (([DBBatidos]![CodGrupo]=7)=False) AND ((DBBatidos.Vigencia)=True) AND ((Cotizaciones.Subproducto)=False) AND ((SubReceta.codporcion) Is Null) AND (([Cotizaciones]![Marca] & " "="Almacen Global ")=False) AND ((Cotizaciones.Prioridad)=True) AND ((Cotizaciones.Descontinuado)=False))
ORDER BY TiposVariables.Orden, DBIngredientes.Nombre;

