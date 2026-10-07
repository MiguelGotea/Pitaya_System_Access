-- ==========================================================
-- Tabla    : SubPorcionamiento
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [SubPorcionamiento] (
    [CodSubPorcionamiento] TYPE_4,
    [Procedencia] TYPE_4,
    [CodProcesamiento] TYPE_4,
    [Cantidad] TYPE_7 NOT NULL,
    [Fecha] TYPE_8,
    [HInicial] TYPE_8,
    [HFinal] TYPE_8,
    [CodOperario] TYPE_4
);

ALTER TABLE [SubPorcionamiento] ADD PRIMARY KEY ([CodSubPorcionamiento]);


