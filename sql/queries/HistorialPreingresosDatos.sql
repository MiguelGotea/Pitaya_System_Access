-- ==========================================================
-- Consulta : HistorialPreingresosDatos
-- Exportado: 2026-10-07 07:17:22
-- ==========================================================

SELECT numerosemana([PreIngresoPitaya]![Fecha]) AS aSemana, nombreproductocoti([SubPreIngresosPitaya]![CodCotizacion]) AS Nombre, PreIngresoPitaya.Fecha, SubPreIngresosPitaya.Cantidad, DBIngredientes.Tipo, DBIngredientes.TIPO1, DBIngredientes.TIPO2, SubPreIngresosPitaya.CodCotizacion, PreIngresoPitaya.Destino
FROM ((DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN SubPreIngresosPitaya ON Cotizaciones.CodCotizacion = SubPreIngresosPitaya.CodCotizacion) INNER JOIN PreIngresoPitaya ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya
WHERE (((numerosemana([PreIngresoPitaya]![Fecha]))=[Formularios]![Menu Gestion]![preingseman]) AND ((SubPreIngresosPitaya.CodCotizacion)<>700 And (SubPreIngresosPitaya.CodCotizacion)<>795 And (SubPreIngresosPitaya.CodCotizacion)<>438 And (SubPreIngresosPitaya.CodCotizacion)<>436 And (SubPreIngresosPitaya.CodCotizacion)<>437 And (SubPreIngresosPitaya.CodCotizacion)<>756 And (SubPreIngresosPitaya.CodCotizacion)<>796) AND ((PreIngresoPitaya.Destino)="Pitaya " & codigolocal()))
ORDER BY PreIngresoPitaya.Fecha;

