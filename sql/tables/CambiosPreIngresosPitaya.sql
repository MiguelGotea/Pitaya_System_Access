-- ==========================================================
-- Tabla    : CambiosPreIngresosPitaya
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [CambiosPreIngresosPitaya] (
    [CodCambiosPreIngresoPitaya] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [Cantidad] TYPE_7 NOT NULL,
    [CodPreIngresoPitaya] TYPE_4,
    [CodSubPreIngresoPitaya] TYPE_4
);

ALTER TABLE [CambiosPreIngresosPitaya] ADD PRIMARY KEY ([CodCambiosPreIngresoPitaya]);


