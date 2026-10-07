-- ==========================================================
-- Tabla    : Grupos
-- Campos   : 16
-- Exportado: 2026-10-07 07:17:26
-- ==========================================================

CREATE TABLE [Grupos] (
    [CodGrupo] TYPE_4,
    [NombreGrupo] VARCHAR(255),
    [Tipo] VARCHAR(255),
    [prioridad] TYPE_4,
    [SumaPuntos] TYPE_1,
    [Imprimible] TYPE_1,
    [Despacho] TYPE_1,
    [Editable] TYPE_1,
    [Tamanos] TYPE_1,
    [Vidrio] TYPE_1,
    [EstacionTrabajo] TYPE_4,
    [Endulzante] TYPE_1,
    [Preparacion] TYPE_1,
    [control] TYPE_1,
    [alias] VARCHAR(255),
    [MenuVendible] TYPE_1
);

ALTER TABLE [Grupos] ADD PRIMARY KEY ([CodGrupo]);


