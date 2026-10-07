-- ==========================================================
-- Tabla    : AnulacionPedidos
-- Campos   : 9
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [AnulacionPedidos] (
    [CodAnulacionPedidos] TYPE_4,
    [CodPedido] TYPE_4,
    [HoraSolicitada] TYPE_8,
    [HoraAnulada] TYPE_8,
    [Status] TYPE_4,
    [Modalidad] TYPE_4,
    [CodPedidoCambio] TYPE_4,
    [Motivo] VARCHAR(255),
    [CodMotivoAnulacion] TYPE_4
);

ALTER TABLE [AnulacionPedidos] ADD PRIMARY KEY ([CodAnulacionPedidos]);


