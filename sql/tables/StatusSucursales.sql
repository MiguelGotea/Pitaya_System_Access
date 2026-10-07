-- ==========================================================
-- Tabla    : StatusSucursales
-- Campos   : 12
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [StatusSucursales] (
    [CodStatusSucursales] TYPE_4,
    [CodLocal] TYPE_4,
    [Activo] TYPE_1,
    [Nombre] VARCHAR(255),
    [Ciudad] VARCHAR(255),
    [Sucursal] TYPE_1,
    [CajaChica] TYPE_4,
    [BotTelegram] VARCHAR(255),
    [PrimerEnvio] VARCHAR(255),
    [SegundoEnvio] VARCHAR(255),
    [Latitude] TYPE_7,
    [Longitude] TYPE_7
);

ALTER TABLE [StatusSucursales] ADD PRIMARY KEY ([CodStatusSucursales]);


