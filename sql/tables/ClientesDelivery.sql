-- ==========================================================
-- Tabla    : ClientesDelivery
-- Campos   : 10
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [ClientesDelivery] (
    [CodClientesDelivery] TYPE_4,
    [CodLocal] TYPE_4,
    [Telefono] TYPE_4,
    [Direccion] VARCHAR(255),
    [TipoPedido] VARCHAR(255),
    [CodProovedoresDelivery] TYPE_4,
    [CostoDelivery] TYPE_4,
    [CodigoMotorizado] TYPE_4,
    [Distancia] TYPE_4,
    [Coordinada] VARCHAR(255)
);

ALTER TABLE [ClientesDelivery] ADD PRIMARY KEY ([CodClientesDelivery]);


