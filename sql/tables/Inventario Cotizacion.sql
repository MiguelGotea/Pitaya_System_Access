-- ==========================================================
-- Tabla    : Inventario Cotizacion
-- Campos   : 10
-- Exportado: 2026-10-07 07:17:27
-- ==========================================================

CREATE TABLE [Inventario Cotizacion] (
    [CodICotizacion] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8 NOT NULL,
    [lista] TYPE_4,
    [CodOperario] TYPE_4,
    [primerenvio] TYPE_4,
    [segundoenvio] TYPE_4,
    [cantidadunidad] TYPE_4,
    [cantidadpaquete] TYPE_4
);

ALTER TABLE [Inventario Cotizacion] ADD PRIMARY KEY ([CodICotizacion]);


