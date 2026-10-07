-- ==========================================================
-- Tabla    : CalParetoSemanal
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [CalParetoSemanal] (
    [CodCalParSem] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Posicion] TYPE_4,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8
);

ALTER TABLE [CalParetoSemanal] ADD PRIMARY KEY ([CodCalParSem]);


