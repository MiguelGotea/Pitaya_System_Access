-- ==========================================================
-- Tabla    : StatusPedidosCentral
-- Campos   : 9
-- Exportado: 2026-10-22 22:23:23
-- ==========================================================

CREATE TABLE [StatusPedidosCentral] (
    [CodStatusPedidosCentral] TYPE_4,
    [CodPedidoCentral] TYPE_4,
    [HoraAprobadoCentral] TYPE_8,
    [AprobadoCentral] TYPE_1,
    [Fecha] TYPE_8,
    [Sucursal] TYPE_4,
    [CodPedidoSucursal] TYPE_4,
    [HoraAprobadoSucursal] TYPE_8,
    [AprobadoSucursal] TYPE_1
);

ALTER TABLE [StatusPedidosCentral] ADD PRIMARY KEY ([CodStatusPedidosCentral]);


