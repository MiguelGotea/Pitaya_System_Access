-- ==========================================================
-- Tabla    : Afiliados
-- Campos   : 6
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

CREATE TABLE [Afiliados] (
    [CodAfiliado] TYPE_4,
    [Nombre] VARCHAR(255),
    [Direccion] VARCHAR(255),
    [Rubro] VARCHAR(255),
    [Ciudad] VARCHAR(255),
    [Vigente] TYPE_1
);

ALTER TABLE [Afiliados] ADD PRIMARY KEY ([CodAfiliado]);


