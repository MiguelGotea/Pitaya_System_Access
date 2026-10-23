-- ==========================================================
-- Tabla    : PreIngresoPitaya
-- Campos   : 6
-- Exportado: 2026-10-22 22:23:22
-- ==========================================================

CREATE TABLE [PreIngresoPitaya] (
    [CodPreIngresoPitaya] TYPE_4,
    [Fecha] TYPE_8,
    [Hora] TYPE_8,
    [Destino] VARCHAR(255),
    [Validado] TYPE_1,
    [Impreso] TYPE_1
);

ALTER TABLE [PreIngresoPitaya] ADD PRIMARY KEY ([CodPreIngresoPitaya]);


