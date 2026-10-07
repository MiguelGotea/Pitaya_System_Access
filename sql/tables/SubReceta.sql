-- ==========================================================
-- Tabla    : SubReceta
-- Campos   : 9
-- Exportado: 2026-10-07 07:17:30
-- ==========================================================

CREATE TABLE [SubReceta] (
    [CodSubReceta] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [CodBatido] VARCHAR(255),
    [Cantidad] TYPE_7,
    [Tipo] VARCHAR(255),
    [codporcion] TYPE_4,
    [InsumoClave] TYPE_1,
    [tiposervido] VARCHAR(255),
    [ordenreceta] TYPE_4
);

ALTER TABLE [SubReceta] ADD PRIMARY KEY ([CodSubReceta]);


