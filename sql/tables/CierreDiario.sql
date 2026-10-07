-- ==========================================================
-- Tabla    : CierreDiario
-- Campos   : 13
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [CierreDiario] (
    [CodigoCierre] TYPE_4,
    [HoraInicial] TYPE_8,
    [HoraFinal] TYPE_8,
    [Fecha] TYPE_8,
    [CodOperario] TYPE_4,
    [MFCor] TYPE_7,
    [MFDol] TYPE_7,
    [Faltante] TYPE_4,
    [TotalHugo] TYPE_7,
    [TotalPedidosYa] TYPE_7,
    [TotalTransferencia] TYPE_7,
    [TotalPOS] TYPE_7,
    [Observaciones] VARCHAR(255)
);

ALTER TABLE [CierreDiario] ADD PRIMARY KEY ([CodigoCierre]);


