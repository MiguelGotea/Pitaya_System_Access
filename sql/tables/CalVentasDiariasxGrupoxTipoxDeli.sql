-- ==========================================================
-- Tabla    : CalVentasDiariasxGrupoxTipoxDeli
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [CalVentasDiariasxGrupoxTipoxDeli] (
    [CodCalVentasDiariasxGrupo] TYPE_4,
    [Fecha] TYPE_8,
    [CodGrupo] TYPE_4,
    [TipoPago] TYPE_3,
    [CodDeliv] TYPE_3,
    [Cantidad] TYPE_7,
    [Monto] TYPE_7,
    [FechaPublicada] TYPE_8
);

ALTER TABLE [CalVentasDiariasxGrupoxTipoxDeli] ADD PRIMARY KEY ([CodCalVentasDiariasxGrupo]);


