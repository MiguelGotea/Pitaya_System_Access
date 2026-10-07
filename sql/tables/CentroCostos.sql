-- ==========================================================
-- Tabla    : CentroCostos
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [CentroCostos] (
    [CodCentroCostos] TYPE_4,
    [Codigo] TYPE_4,
    [CodigoTexto] VARCHAR(255),
    [Nombre] VARCHAR(255),
    [Clase] VARCHAR(255),
    [Grupo] TYPE_4,
    [Activo] TYPE_1,
    [Sector] VARCHAR(255)
);

ALTER TABLE [CentroCostos] ADD PRIMARY KEY ([CodCentroCostos]);


