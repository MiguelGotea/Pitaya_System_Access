-- ==========================================================
-- Tabla    : tClientesClub
-- Campos   : 11
-- Exportado: 2026-10-07 07:17:30
-- ==========================================================

CREATE TABLE [tClientesClub] (
    [CodCliente] TYPE_4 NOT NULL,
    [Nombre] VARCHAR(255),
    [Apellidos] VARCHAR(255),
    [Celular] VARCHAR(255),
    [Cumpleanos] TYPE_8,
    [Correo] VARCHAR(255),
    [Fecha de Inscripcion] TYPE_8,
    [PuntosIniciales] TYPE_4,
    [local] VARCHAR(255),
    [Genero] VARCHAR(255),
    [CodAfiliado] TYPE_4
);

ALTER TABLE [tClientesClub] ADD PRIMARY KEY ([CodCliente]);


