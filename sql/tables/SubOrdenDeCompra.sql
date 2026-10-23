-- ==========================================================
-- Tabla    : SubOrdenDeCompra
-- Campos   : 8
-- Exportado: 2026-10-22 22:23:23
-- ==========================================================

CREATE TABLE [SubOrdenDeCompra] (
    [codsubordendecompra] TYPE_4,
    [codordendecompra] TYPE_4,
    [codcotizacion] TYPE_4,
    [cantidadorden] TYPE_7,
    [destino] TYPE_4,
    [costounitario] TYPE_7,
    [aplicaiva] TYPE_1,
    [Observaciones] VARCHAR(255)
);

ALTER TABLE [SubOrdenDeCompra] ADD PRIMARY KEY ([codsubordendecompra]);


