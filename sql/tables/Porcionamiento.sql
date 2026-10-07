-- ==========================================================
-- Tabla    : Porcionamiento
-- Campos   : 11
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [Porcionamiento] (
    [CodPorcionamiento] TYPE_4,
    [CodCotizacion] TYPE_4 NOT NULL,
    [CodProcesamiento] TYPE_4 NOT NULL,
    [Cantidad] TYPE_4 NOT NULL,
    [Observaciones] VARCHAR(255),
    [Fecha] TYPE_8 NOT NULL,
    [CodOperario] TYPE_4 NOT NULL,
    [Procedencia] TYPE_4,
    [CodSubPorcionamiento] TYPE_4,
    [HInicial] TYPE_8,
    [HFinal] TYPE_8
);

ALTER TABLE [Porcionamiento] ADD PRIMARY KEY ([CodPorcionamiento]);


