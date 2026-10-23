-- ==========================================================
-- Tabla    : CalConversionSemanal
-- Campos   : 5
-- Exportado: 2026-10-22 22:23:15
-- ==========================================================

CREATE TABLE [CalConversionSemanal] (
    [CodCalConvSem] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Conversion] TYPE_7,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8
);

ALTER TABLE [CalConversionSemanal] ADD PRIMARY KEY ([CodCalConvSem]);


