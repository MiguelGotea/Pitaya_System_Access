-- ==========================================================
-- Tabla    : CambiosPreIngresosPitaya
-- Campos   : 4
-- Exportado: 2026-10-14 21:56:14
-- ==========================================================

CREATE TABLE [CambiosPreIngresosPitaya] (
    [CodCambiosPreIngresoPitaya] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [Cantidad] TYPE_7 NOT NULL,
    [CodPreIngresoPitaya] TYPE_4
);

ALTER TABLE [CambiosPreIngresosPitaya] ADD PRIMARY KEY ([CodCambiosPreIngresoPitaya]);


