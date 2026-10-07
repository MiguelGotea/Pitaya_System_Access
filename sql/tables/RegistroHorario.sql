-- ==========================================================
-- Tabla    : RegistroHorario
-- Campos   : 6
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [RegistroHorario] (
    [CodHorario] TYPE_4,
    [Ingreso] TYPE_8,
    [Salida] TYPE_8,
    [Fecha] TYPE_8,
    [CodOperario] TYPE_4,
    [Cancelado] TYPE_1
);

ALTER TABLE [RegistroHorario] ADD PRIMARY KEY ([CodHorario]);


