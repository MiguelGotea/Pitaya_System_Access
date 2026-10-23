-- ==========================================================
-- Tabla    : CalConsumoNoPorcionSemanal
-- Campos   : 6
-- Exportado: 2026-10-22 22:23:15
-- ==========================================================

CREATE TABLE [CalConsumoNoPorcionSemanal] (
    [CodCalConsuSem] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Consumo] TYPE_7,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8,
    [ConsumoMaximo] TYPE_7
);

ALTER TABLE [CalConsumoNoPorcionSemanal] ADD PRIMARY KEY ([CodCalConsuSem]);


