-- ==========================================================
-- Tabla    : NotaDePedido
-- Campos   : 27
-- Exportado: 2026-10-07 07:17:28
-- ==========================================================

CREATE TABLE [NotaDePedido] (
    [CodPedido] TYPE_4,
    [Fecha] TYPE_8,
    [Hora] TYPE_8,
    [Modalidad] TYPE_1,
    [CodCliente] TYPE_4,
    [Nombre Provicional] VARCHAR(255),
    [senuelo] VARCHAR(255),
    [Anulado] TYPE_1,
    [MotivoAnulado] VARCHAR(255),
    [POS] TYPE_1,
    [Impreso] TYPE_1,
    [Delivery] TYPE_4,
    [TotalGuardado] TYPE_7,
    [Transferencia] TYPE_1,
    [DeliveryComision] TYPE_1,
    [Impresiones] TYPE_4,
    [CodMotorizado] TYPE_4,
    [HoraCreado] TYPE_8,
    [HoraImpreso] TYPE_8,
    [HoraIngresoProducto] TYPE_8,
    [TipoAnulado] TYPE_4,
    [ProductosEliminados] TYPE_1,
    [Propina] TYPE_7,
    [Observaciones] VARCHAR(255),
    [CodClientesDelivery] TYPE_4,
    [recibecor] TYPE_4,
    [recibedol] TYPE_4
);

ALTER TABLE [NotaDePedido] ADD PRIMARY KEY ([CodPedido]);


