-- ==========================================================
-- Tabla    : DBIngredientes
-- Campos   : 25
-- Exportado: 2026-10-07 07:17:26
-- ==========================================================

CREATE TABLE [DBIngredientes] (
    [CodIngrediente] VARCHAR(255) NOT NULL,
    [Nombre] VARCHAR(255),
    [NombreSinProcesar] VARCHAR(255),
    [NombreProcesado] VARCHAR(255),
    [NombreReceta] VARCHAR(255),
    [Unidad] VARCHAR(255),
    [OrdenListaControles] TYPE_4,
    [ListaImportante] TYPE_1,
    [Inventario] TYPE_1,
    [Tipo] VARCHAR(255),
    [TIPO1] VARCHAR(255),
    [TIPO2] VARCHAR(255),
    [compralocal] TYPE_1,
    [Consumible] TYPE_1,
    [Cocina] TYPE_1,
    [ComprasSemana] TYPE_7,
    [Vigente] TYPE_1,
    [ProovedorPrincipal] TYPE_4,
    [ProovedorSecundario] TYPE_4,
    [ConversionGramos] TYPE_7,
    [CUEstandar] TYPE_7,
    [CodCeCoSubCuentas] TYPE_4,
    [CodControlAlmacenes] TYPE_4,
    [presentacionpreparacion] VARCHAR(255),
    [conversionpreparacion] TYPE_7
);

ALTER TABLE [DBIngredientes] ADD PRIMARY KEY ([CodIngrediente]);


