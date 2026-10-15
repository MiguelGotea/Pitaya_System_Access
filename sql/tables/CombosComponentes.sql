-- ==========================================================
-- Tabla    : CombosComponentes
-- Campos   : 6
-- Exportado: 2026-10-14 21:56:14
-- ==========================================================

CREATE TABLE [CombosComponentes] (
    [CodComboComponente] TYPE_4,
    [CodBatido] VARCHAR(255),
    [GrupoComponente] TYPE_7,
    [ProductoEspecifico] VARCHAR(255),
    [Menu] VARCHAR(255),
    [Cantidad] TYPE_4
);

ALTER TABLE [CombosComponentes] ADD PRIMARY KEY ([CodComboComponente]);


