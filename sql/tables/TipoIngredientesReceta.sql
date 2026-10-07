-- ==========================================================
-- Tabla    : TipoIngredientesReceta
-- Campos   : 4
-- Exportado: 2026-10-07 07:17:30
-- ==========================================================

CREATE TABLE [TipoIngredientesReceta] (
    [CodTipoIngredientesReceta] VARCHAR(255) NOT NULL,
    [Nombre] VARCHAR(255),
    [Orden] TYPE_4,
    [GrupoTipoReceta] VARCHAR(255)
);

ALTER TABLE [TipoIngredientesReceta] ADD PRIMARY KEY ([CodTipoIngredientesReceta]);


