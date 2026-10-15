-- ==========================================================
-- Tabla    : CalConsumoPorcionSemanal
-- Campos   : 6
-- Exportado: 2026-10-14 21:56:13
-- ==========================================================

CREATE TABLE [CalConsumoPorcionSemanal] (
    [CodCalPorcionSem] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Consumo] TYPE_7,
    [Semana] TYPE_4,
    [FechaPublicada] TYPE_8,
    [ConsumoMaximo] TYPE_7
);

ALTER TABLE [CalConsumoPorcionSemanal] ADD PRIMARY KEY ([CodCalPorcionSem]);


