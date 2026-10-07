-- ==========================================================
-- Tabla    : Inventario Cotizacion Temporal
-- Campos   : 10
-- Exportado: 2026-10-07 07:17:27
-- ==========================================================

CREATE TABLE [Inventario Cotizacion Temporal] (
    [CodICotizacion] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8 NOT NULL,
    [lista] TYPE_4,
    [CodOperario] TYPE_4,
    [primerenvio] TYPE_4,
    [segundoenvio] TYPE_4,
    [cantidadunidad] TYPE_7,
    [cantidadpaquete] TYPE_7
);

ALTER TABLE [Inventario Cotizacion Temporal] ADD PRIMARY KEY ([CodICotizacion]);


