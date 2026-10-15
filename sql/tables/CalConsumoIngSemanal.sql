-- ==========================================================
-- Tabla    : CalConsumoIngSemanal
-- Campos   : 6
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

CREATE TABLE [CalConsumoIngSemanal] (
    [CodCalConsuSem] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Consumo] TYPE_7,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8,
    [ConsumoMaximo] TYPE_7
);

ALTER TABLE [CalConsumoIngSemanal] ADD PRIMARY KEY ([CodCalConsuSem]);


