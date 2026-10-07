-- ==========================================================
-- Tabla    : HorariosSucursalesAprobados
-- Campos   : 48
-- Exportado: 2026-10-07 07:17:27
-- ==========================================================

CREATE TABLE [HorariosSucursalesAprobados] (
    [CodHorariosSucursales] TYPE_4,
    [CodOperario] TYPE_4,
    [CodTipoLunes] TYPE_4,
    [TE1Lunes] TYPE_8,
    [TE2Lunes] TYPE_8,
    [TS1Lunes] TYPE_8,
    [TS2Lunes] TYPE_8,
    [CodTipoMartes] TYPE_4,
    [TE1Martes] TYPE_8,
    [TE2Martes] TYPE_8,
    [TS1Martes] TYPE_8,
    [TS2Martes] TYPE_8,
    [CodTipoMiercoles] TYPE_4,
    [TE1Miercoles] TYPE_8,
    [TE2Miercoles] TYPE_8,
    [TS1Miercoles] TYPE_8,
    [TS2Miercoles] TYPE_8,
    [CodTipoJueves] TYPE_4,
    [TE1Jueves] TYPE_8,
    [TE2Jueves] TYPE_8,
    [TS1Jueves] TYPE_8,
    [TS2Jueves] TYPE_8,
    [CodTipoViernes] TYPE_4,
    [TE1Viernes] TYPE_8,
    [TE2Viernes] TYPE_8,
    [TS1Viernes] TYPE_8,
    [TS2Viernes] TYPE_8,
    [CodTipoSabado] TYPE_4,
    [TE1Sabado] TYPE_8,
    [TE2Sabado] TYPE_8,
    [TS1Sabado] TYPE_8,
    [TS2Sabado] TYPE_8,
    [CodTipoDomingo] TYPE_4,
    [TE1Domingo] TYPE_8,
    [TE2Domingo] TYPE_8,
    [TS1Domingo] TYPE_8,
    [TS2Domingo] TYPE_8,
    [semana] TYPE_4,
    [HorasFaltantes] TYPE_7,
    [Sucursal] TYPE_4,
    [Aprobado] TYPE_8,
    [NotaLunes] VARCHAR(255),
    [NotaMartes] VARCHAR(255),
    [NotaMiercoles] VARCHAR(255),
    [NotaJueves] VARCHAR(255),
    [NotaViernes] VARCHAR(255),
    [NotaSabado] VARCHAR(255),
    [NotaDomingo] VARCHAR(255)
);

ALTER TABLE [HorariosSucursalesAprobados] ADD PRIMARY KEY ([CodHorariosSucursales]);


