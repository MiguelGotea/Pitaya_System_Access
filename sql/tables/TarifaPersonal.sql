-- ==========================================================
-- Tabla    : TarifaPersonal
-- Campos   : 14
-- Exportado: 2026-10-07 07:17:30
-- ==========================================================

CREATE TABLE [TarifaPersonal] (
    [CodTarifa] TYPE_4,
    [CodOperario] TYPE_4,
    [FechaInicial] TYPE_8,
    [FechaFinal] TYPE_8,
    [Tarifa] TYPE_7,
    [HL] TYPE_7,
    [HMA] TYPE_7,
    [HMI] TYPE_7,
    [HJ] TYPE_7,
    [HV] TYPE_7,
    [HS] TYPE_7,
    [HD] TYPE_7,
    [Fijo] TYPE_1,
    [SalarioFijo] TYPE_4
);

ALTER TABLE [TarifaPersonal] ADD PRIMARY KEY ([CodTarifa]);


