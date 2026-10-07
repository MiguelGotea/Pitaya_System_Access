-- ==========================================================
-- Tabla    : Inventario Ingrediente
-- Campos   : 6
-- Exportado: 2026-10-07 07:17:27
-- ==========================================================

CREATE TABLE [Inventario Ingrediente] (
    [CodIIngrediente] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8,
    [lista] TYPE_4,
    [CodOperario] TYPE_4
);

ALTER TABLE [Inventario Ingrediente] ADD PRIMARY KEY ([CodIIngrediente]);


