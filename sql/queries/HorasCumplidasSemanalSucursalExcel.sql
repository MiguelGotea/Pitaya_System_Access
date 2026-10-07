-- ==========================================================
-- Consulta : HorasCumplidasSemanalSucursalExcel
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT Operarios.CodOperario, NombreOperario([Operarios]![CodOperario]) AS Nombre, numerosemana([FechaSistema]![Dates]) AS sema, MaxMinSemana(numerosemana([FechaSistema]![Dates]),0) AS Desde, MaxMinSemana(numerosemana([FechaSistema]![Dates]),1) AS Hasta, HorasSemanalesCumplidasSinRedondeo(numerosemana([FechaSistema]![Dates]),[Operarios]![CodOperario]) AS Horas, [FechaSistema]![Dates] Between [Formularios]![Pago Personal]![desde] And [Formularios]![Pago Personal]![hasta] AS Expr2, Operarios.Operativo, Operarios.Sucursal
FROM FechaSistema, Operarios
GROUP BY Operarios.CodOperario, NombreOperario([Operarios]![CodOperario]), numerosemana([FechaSistema]![Dates]), MaxMinSemana(numerosemana([FechaSistema]![Dates]),0), MaxMinSemana(numerosemana([FechaSistema]![Dates]),1), HorasSemanalesCumplidasSinRedondeo(numerosemana([FechaSistema]![Dates]),[Operarios]![CodOperario]), [FechaSistema]![Dates] Between [Formularios]![Pago Personal]![desde] And [Formularios]![Pago Personal]![hasta], Operarios.Operativo, Operarios.Sucursal
HAVING ((([FechaSistema]![Dates] Between [Formularios]![Pago Personal]![desde] And [Formularios]![Pago Personal]![hasta])=-1) AND ((Operarios.Operativo)=True) AND ((Operarios.Sucursal)=codigolocal()));

