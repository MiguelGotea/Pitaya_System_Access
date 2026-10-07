-- ==========================================================
-- Tabla    : DBBatidos
-- Campos   : 12
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [DBBatidos] (
    [CodBatido] VARCHAR(255) NOT NULL,
    [Nombre] VARCHAR(255),
    [CodGrupo] TYPE_4,
    [Medida] VARCHAR(255),
    [Precio] TYPE_4,
    [Vigencia] TYPE_1,
    [CodigoBarras] VARCHAR(255),
    [Marca] VARCHAR(255),
    [CompraVenta] TYPE_1,
    [CodSubGrupo] TYPE_4,
    [Endulzante] TYPE_4,
    [Desde] TYPE_8
);

ALTER TABLE [DBBatidos] ADD PRIMARY KEY ([CodBatido]);


