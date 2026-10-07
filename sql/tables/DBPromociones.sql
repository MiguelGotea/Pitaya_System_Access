-- ==========================================================
-- Tabla    : DBPromociones
-- Campos   : 16
-- Exportado: 2026-10-07 07:17:26
-- ==========================================================

CREATE TABLE [DBPromociones] (
    [CodPromocion] TYPE_4,
    [Nombre] VARCHAR(255),
    [RatioDescuento] TYPE_7,
    [DescuentoFijo] TYPE_7,
    [Condiciones] TYPE_12,
    [Vigencia] VARCHAR(255),
    [Cliente] VARCHAR(255),
    [Activo] TYPE_1,
    [Local] TYPE_4,
    [diaespecifico] TYPE_4,
    [usointerno] TYPE_1,
    [dia] VARCHAR(255),
    [hora] VARCHAR(255),
    [aplicaen] TYPE_12,
    [mododeaplicacion] TYPE_12,
    [tipocliente] VARCHAR(255)
);

ALTER TABLE [DBPromociones] ADD PRIMARY KEY ([CodPromocion]);


