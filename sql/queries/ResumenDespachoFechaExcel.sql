-- ==========================================================
-- Consulta : ResumenDespachoFechaExcel
-- Exportado: 2026-10-07 06:21:34
-- ==========================================================
SELECT PreIngresoPitaya.Fecha, SubPreIngresosPitaya.CodCotizacion, Cotizaciones.CodIngrediente, Right([PreIngresoPitaya]![Destino],Len([PreIngresoPitaya]![Destino])-7) AS [local], DBIngredientes.TIPO2, DBIngredientes.Tipo, nombreproductocotiprocesado([SubPreIngresosPitaya]![CodCotizacion]) AS Producto, SubPreIngresosPitaya.Cantidad, numerosemana([PreIngresoPitaya]![Fecha]) AS semana, PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion]) AS mezcla, Cotizaciones.Conversion
FROM ((SubPreIngresosPitaya INNER JOIN PreIngresoPitaya ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya) INNER JOIN Cotizaciones ON SubPreIngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion) INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente
WHERE (((PreIngresoPitaya.Fecha) Between [Formularios]![Menu Gestion]![despainicio] And [Formularios]![Menu Gestion]![despafinal]) AND ((PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion]))=0))
ORDER BY PreIngresoPitaya.Fecha, Right([PreIngresoPitaya]![Destino],Len([PreIngresoPitaya]![Destino])-7), DBIngredientes.TIPO2, DBIngredientes.Tipo, nombreproductocotiprocesado([SubPreIngresosPitaya]![CodCotizacion]);

