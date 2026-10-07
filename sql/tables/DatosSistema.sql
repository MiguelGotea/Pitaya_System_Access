-- ==========================================================
-- Tabla    : DatosSistema
-- Campos   : 17
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [DatosSistema] (
    [Nombre] VARCHAR(255),
    [Direccion] VARCHAR(255),
    [Codigo] VARCHAR(255) NOT NULL,
    [Ciudad] VARCHAR(255),
    [clavewifi] VARCHAR(255),
    [CodSistema] TYPE_4,
    [Teamviewer] TYPE_4,
    [VolumenMusica] TYPE_4,
    [IngresoInsumos] TYPE_4,
    [ConcursoActivo] TYPE_1,
    [CanalCaja] VARCHAR(255),
    [IpCamara] VARCHAR(255),
    [ClaveCamara] VARCHAR(255),
    [ActivarDelivery] TYPE_1,
    [IdBotTe] VARCHAR(255),
    [IdGrupoTel] VARCHAR(255),
    [Activo] TYPE_1
);

ALTER TABLE [DatosSistema] ADD PRIMARY KEY ([Codigo]);


