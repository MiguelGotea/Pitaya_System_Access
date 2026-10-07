-- ==========================================================
-- Tabla    : Procesamiento
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [Procesamiento] (
    [CodProcesamiento] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7 NOT NULL,
    [MedidaInicial] TYPE_7,
    [MedidaFinal] TYPE_7 NOT NULL,
    [Fecha] TYPE_8,
    [Observaciones] VARCHAR(255),
    [Operario] TYPE_4 NOT NULL
);

ALTER TABLE [Procesamiento] ADD PRIMARY KEY ([CodProcesamiento]);


