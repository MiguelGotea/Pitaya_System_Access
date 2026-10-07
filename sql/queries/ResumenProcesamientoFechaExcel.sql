-- ==========================================================
-- Consulta : ResumenProcesamientoFechaExcel
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT SubPorcionamiento.Fecha, nombreproductocotiprocesado([SubPorcionamiento]![Procedencia]) AS Producto, SubPorcionamiento.Cantidad, NombreOperario([SubPorcionamiento]![CodOperario]) AS Personal, Round(([SubPorcionamiento]![HFinal]-[SubPorcionamiento]![HInicial])*24*60,0) AS Tiempo, SubPorcionamiento.HInicial, SubPorcionamiento.HFinal, DLookUp("[Conversion]","[Cotizaciones]","[CodCotizacion]=" & [SubPorcionamiento]![Procedencia]) AS condi
FROM SubPorcionamiento
WHERE (((SubPorcionamiento.Fecha) Between [Formularios]![Menu Gestion]![dproduc] And [Formularios]![Menu Gestion]![hprodu]) AND ((DLookUp("[Conversion]","[Cotizaciones]","[CodCotizacion]=" & [SubPorcionamiento]![Procedencia]))=0));

