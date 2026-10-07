-- ==========================================================
-- Tabla    : IngresosPitaya
-- Campos   : 6
-- Exportado: 2026-10-07 07:17:27
-- ==========================================================

CREATE TABLE [IngresosPitaya] (
    [CodIngresoPitaya] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [Cantidad] TYPE_7 NOT NULL,
    [Fecha] TYPE_8 NOT NULL,
    [Procedencia] TYPE_4,
    [TipoProcedencia] VARCHAR(255)
);

ALTER TABLE [IngresosPitaya] ADD PRIMARY KEY ([CodIngresoPitaya]);


