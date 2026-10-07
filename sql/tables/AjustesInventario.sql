-- ==========================================================
-- Tabla    : AjustesInventario
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

CREATE TABLE [AjustesInventario] (
    [CodAjustesInventario] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8,
    [Observacion] VARCHAR(255)
);

ALTER TABLE [AjustesInventario] ADD PRIMARY KEY ([CodAjustesInventario]);


