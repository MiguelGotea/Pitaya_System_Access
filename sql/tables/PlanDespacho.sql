-- ==========================================================
-- Tabla    : PlanDespacho
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:28
-- ==========================================================

CREATE TABLE [PlanDespacho] (
    [CodPlanDespacho] TYPE_4,
    [Local] TYPE_4,
    [Lista] TYPE_4,
    [CodAlmacenamiento] TYPE_4,
    [Fecha] TYPE_4,
    [Carga] TYPE_7,
    [SemanaMes] VARCHAR(255),
    [Limitado] TYPE_1
);

ALTER TABLE [PlanDespacho] ADD PRIMARY KEY ([CodPlanDespacho]);


