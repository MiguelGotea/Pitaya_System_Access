-- ==========================================================
-- Tabla    : Proovedores
-- Campos   : 14
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [Proovedores] (
    [CodProovedor] TYPE_4,
    [Nombre] VARCHAR(255),
    [Numero] VARCHAR(255),
    [Direccion] VARCHAR(255),
    [CompraExclusivaSucursal] TYPE_4,
    [Directo] TYPE_1,
    [Contacto] VARCHAR(255),
    [CompraProgramada] TYPE_1,
    [CBNumero] VARCHAR(255),
    [CBTitular] VARCHAR(255),
    [CBMoneda] VARCHAR(255),
    [CBBanco] VARCHAR(255),
    [Modalidad] VARCHAR(255),
    [TipoDePago] VARCHAR(255)
);

ALTER TABLE [Proovedores] ADD PRIMARY KEY ([CodProovedor]);


