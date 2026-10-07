-- ==========================================================
-- Tabla    : FiltroMostradorVigentes
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:26
-- ==========================================================

CREATE TABLE [FiltroMostradorVigentes] (
    [CodFiltroNoPorcionesVigentes] TYPE_4,
    [Tipo] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [CodCotizacion] TYPE_4,
    [FechaActualizacion] TYPE_8
);

ALTER TABLE [FiltroMostradorVigentes] ADD PRIMARY KEY ([CodFiltroNoPorcionesVigentes]);


