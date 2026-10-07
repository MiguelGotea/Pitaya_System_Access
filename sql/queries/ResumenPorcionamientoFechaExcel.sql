-- ==========================================================
-- Consulta : ResumenPorcionamientoFechaExcel
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT Porcionamiento.Fecha, nombreproductocotiprocesado([Porcionamiento]![CodCotizacion]) AS Producto, Porcionamiento.Cantidad, NombreOperario([Porcionamiento]![CodOperario]) AS Personal, Round(([Porcionamiento]![HFinal]-[Porcionamiento]![HInicial])*24*60,0) AS Tiempo, Porcionamiento.HInicial, Porcionamiento.HFinal
FROM Porcionamiento
WHERE (((Porcionamiento.Fecha) Between [Formularios]![Menu Gestion]![dproduc] And [Formularios]![Menu Gestion]![hprodu]));

