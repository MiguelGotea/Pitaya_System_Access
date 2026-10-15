-- ==========================================================
-- Tabla    : CompensacionFeriados
-- Campos   : 5
-- Exportado: 2026-10-14 21:56:14
-- ==========================================================

CREATE TABLE [CompensacionFeriados] (
    [CodCompensacionFeriados] TYPE_4,
    [CodOperario] TYPE_4,
    [Inicio] TYPE_8,
    [Final] TYPE_8,
    [HorasPromedio] TYPE_4
);

ALTER TABLE [CompensacionFeriados] ADD PRIMARY KEY ([CodCompensacionFeriados]);


