-- ==========================================================
-- Tabla    : SistemaGlobal
-- Campos   : 17
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [SistemaGlobal] (
    [CodigoSistemaGlobal] TYPE_4,
    [TelegramGerencia] VARCHAR(255),
    [TelegramOperaciones] VARCHAR(255),
    [TelegramAnulaciones] VARCHAR(255),
    [TelegramContabilidad] VARCHAR(255),
    [FormularioCanceladoPedidosYa] VARCHAR(255),
    [ImpresoraTermicaSucursal] VARCHAR(255),
    [ImpresoraTermicaEtiquetas] VARCHAR(255),
    [ImpresoraTintaOficinas] VARCHAR(255),
    [ImpresoraConvertirImagen] VARCHAR(255),
    [TelegramPedidosCentral] VARCHAR(255),
    [DriverHostinger] VARCHAR(255),
    [ServidorHostinger] VARCHAR(255),
    [DBPrincipalHostinger] VARCHAR(255),
    [UsuarioPrincipalHostinger] VARCHAR(255),
    [ClavePrincipalHostinger] VARCHAR(255),
    [GoogleMapsAPI] VARCHAR(255)
);

ALTER TABLE [SistemaGlobal] ADD PRIMARY KEY ([CodigoSistemaGlobal]);


