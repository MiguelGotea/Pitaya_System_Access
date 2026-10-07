-- ==========================================================
-- Tabla    : CalCUIngSemanal
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [CalCUIngSemanal] (
    [CodCalCUIngSem] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [CU] TYPE_7,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8
);

ALTER TABLE [CalCUIngSemanal] ADD PRIMARY KEY ([CodCalCUIngSem]);


