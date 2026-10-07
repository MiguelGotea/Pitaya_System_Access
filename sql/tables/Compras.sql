-- ==========================================================
-- Tabla    : Compras
-- Campos   : 16
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [Compras] (
    [CodIngresoAlmacen] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [Cantidad] TYPE_7 NOT NULL,
    [Fecha] TYPE_8 NOT NULL,
    [CostoTotal] TYPE_7 NOT NULL,
    [Observaciones] VARCHAR(255),
    [CodProveedor] TYPE_4 NOT NULL,
    [Destino] VARCHAR(255),
    [Tipo] VARCHAR(255),
    [Pagado] TYPE_8,
    [NumeroFactura] VARCHAR(255),
    [CodOperario] TYPE_4,
    [Ingresado] TYPE_1,
    [Lote] TYPE_7,
    [Peso] TYPE_7,
    [Hora] TYPE_8
);

ALTER TABLE [Compras] ADD PRIMARY KEY ([CodIngresoAlmacen]);


