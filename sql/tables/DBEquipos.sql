-- ==========================================================
-- Tabla    : DBEquipos
-- Campos   : 11
-- Exportado: 2026-10-14 21:56:14
-- ==========================================================

CREATE TABLE [DBEquipos] (
    [CodEquipos] VARCHAR(255),
    [CodigoRotulado] VARCHAR(255),
    [CodTipoEquipos] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Marca] VARCHAR(255),
    [Modelo] VARCHAR(255),
    [Caracteristicas] VARCHAR(255),
    [Serial] VARCHAR(255),
    [FechaCompra] TYPE_8,
    [Proovedor] VARCHAR(255),
    [Garantia] TYPE_4
);

ALTER TABLE [DBEquipos] ADD PRIMARY KEY ([CodEquipos]);


