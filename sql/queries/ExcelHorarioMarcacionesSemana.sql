-- ==========================================================
-- Consulta : ExcelHorarioMarcacionesSemana
-- Exportado: 2026-10-07 06:21:33
-- ==========================================================
SELECT [Formularios]![Menu Gestion]![semanahorario] AS Semana, Operarios.CodOperario, NombreOperario([Operarios]![CodOperario]) AS Colaborador, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],1)) AS Lunes, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],2)) AS Martes, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],3)) AS Miercoles, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],4)) AS Jueves, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],5)) AS Viernes, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],6)) AS Sabado, horasmarcadastextototal([Operarios]![CodOperario],FechaDeNumeroDiaSemana([Formularios]![Menu Gestion]![semanahorario],7)) AS Domingo, Round(HorasSemanalesCumplidasSinRedondeo([Formularios]![Menu Gestion]![semanahorario],[Operarios]![CodOperario]),1) AS Horas
FROM Operarios
WHERE (((Round(HorasSemanalesCumplidasSinRedondeo([Formularios]![Menu Gestion]![semanahorario],[Operarios]![CodOperario]),1))<>0));

