-- ==========================================================
-- Tabla    : Merma Cotizacion
-- Campos   : 7
-- Exportado: 2026-10-22 22:23:20
-- ==========================================================

CREATE TABLE [Merma Cotizacion] (
    [CodMermaUnidad] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8,
    [Observacion] VARCHAR(255),
    [CodIncidencia] TYPE_4,
    [Operario] TYPE_4
);

ALTER TABLE [Merma Cotizacion] ADD PRIMARY KEY ([CodMermaUnidad]);


