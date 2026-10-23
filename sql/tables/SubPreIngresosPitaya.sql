-- ==========================================================
-- Tabla    : SubPreIngresosPitaya
-- Campos   : 5
-- Exportado: 2026-10-22 22:23:23
-- ==========================================================

CREATE TABLE [SubPreIngresosPitaya] (
    [CodSubPreIngresoPitaya] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [Cantidad] TYPE_7 NOT NULL,
    [CodPreIngresoPitaya] TYPE_4,
    [alerta] TYPE_1
);

ALTER TABLE [SubPreIngresosPitaya] ADD PRIMARY KEY ([CodSubPreIngresoPitaya]);


