-- ==========================================================
-- Tabla    : Cotizaciones
-- Campos   : 20
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [Cotizaciones] (
    [CodCotizacion] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Marca] VARCHAR(255),
    [Linea] VARCHAR(255),
    [Capacidad] VARCHAR(255),
    [Conversion] TYPE_7,
    [Unidad] VARCHAR(255),
    [Prioridad] TYPE_1,
    [Subproducto] TYPE_1,
    [PaquetePorciones] TYPE_4,
    [TiempoOperativo] TYPE_7,
    [CodAlmacenamiento] TYPE_4,
    [MezclaPorcion] TYPE_1,
    [ConversionEstandar] TYPE_7,
    [PresentacionCompra] TYPE_7,
    [Descontinuado] TYPE_1,
    [CompraDirectaSucursal] TYPE_1,
    [OrdenInventario] TYPE_4,
    [Especificaciones] VARCHAR(255),
    [CodigoAlmacenDespacho] VARCHAR(255)
);

ALTER TABLE [Cotizaciones] ADD PRIMARY KEY ([CodCotizacion]);


