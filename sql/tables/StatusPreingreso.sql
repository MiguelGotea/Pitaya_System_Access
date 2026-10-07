-- ==========================================================
-- Tabla    : StatusPreingreso
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [StatusPreingreso] (
    [CodStatusPreingreso] TYPE_4,
    [CodPreIngresoPitaya] TYPE_4,
    [Status] TYPE_1,
    [CodOperario] TYPE_7 NOT NULL,
    [Fecha] TYPE_8
);

ALTER TABLE [StatusPreingreso] ADD PRIMARY KEY ([CodStatusPreingreso]);


