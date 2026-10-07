-- ==========================================================
-- Tabla    : Merma Unidad
-- Campos   : 7
-- Exportado: 2026-10-07 07:17:28
-- ==========================================================

CREATE TABLE [Merma Unidad] (
    [CodMermaUnidad] TYPE_4,
    [CodIngrediente] VARCHAR(255),
    [Cantidad] TYPE_7,
    [Fecha] TYPE_8,
    [Observacion] VARCHAR(255),
    [CodIncidencia] TYPE_4,
    [Operario] TYPE_4
);

ALTER TABLE [Merma Unidad] ADD PRIMARY KEY ([CodMermaUnidad]);


