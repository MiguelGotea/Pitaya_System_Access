-- ==========================================================
-- Tabla    : SubPedido
-- Campos   : 15
-- Exportado: 2026-10-07 07:17:29
-- ==========================================================

CREATE TABLE [SubPedido] (
    [CodSubPedido] TYPE_4,
    [CodBatido] VARCHAR(255),
    [Cantidad] TYPE_7 NOT NULL,
    [CodPedido] TYPE_4,
    [CodPromocion] TYPE_4,
    [Observaciones] VARCHAR(255),
    [Eliminado] TYPE_1,
    [CodCortesia] TYPE_4,
    [SinAzucar] TYPE_1,
    [Azucar] VARCHAR(255),
    [Vinculo] TYPE_4,
    [Empaque] TYPE_1,
    [Puntos] TYPE_7,
    [Monto] TYPE_7,
    [DetallesAdicionales] VARCHAR(255)
);

ALTER TABLE [SubPedido] ADD PRIMARY KEY ([CodSubPedido]);


