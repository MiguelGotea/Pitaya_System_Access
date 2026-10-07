-- ==========================================================
-- Tabla    : ComposicionIngredientes
-- Campos   : 12
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [ComposicionIngredientes] (
    [CodComposicion] TYPE_4,
    [CodIngrediente] VARCHAR(255) NOT NULL,
    [Nombre] VARCHAR(255),
    [Unidad] VARCHAR(255),
    [EnergiaKCAL] TYPE_7,
    [ProteinasGR] TYPE_7,
    [GrasasGR] TYPE_7,
    [CarbohidratosGR] TYPE_7,
    [FibraGR] TYPE_7,
    [AzucarGR] TYPE_7,
    [VitCMG] TYPE_7,
    [Gluten] TYPE_4
);

ALTER TABLE [ComposicionIngredientes] ADD PRIMARY KEY ([CodComposicion]);


