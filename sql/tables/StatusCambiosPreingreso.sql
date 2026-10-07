-- ==========================================================
-- Tabla    : StatusCambiosPreingreso
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [StatusCambiosPreingreso] (
    [CodStatusCambiosPreingreso] TYPE_4,
    [CodCambiosPreIngresoPitaya] TYPE_4,
    [Status] TYPE_1,
    [CodOperario] TYPE_7 NOT NULL,
    [Fecha] TYPE_8
);

ALTER TABLE [StatusCambiosPreingreso] ADD PRIMARY KEY ([CodStatusCambiosPreingreso]);


