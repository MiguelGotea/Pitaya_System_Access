-- ==========================================================
-- Tabla    : OrdenDeCompra
-- Campos   : 5
-- Exportado: 2026-10-22 22:23:21
-- ==========================================================

CREATE TABLE [OrdenDeCompra] (
    [codordendecompra] TYPE_4,
    [codproovedor] TYPE_4,
    [fechaorden] TYPE_8,
    [tipopago] VARCHAR(255),
    [resumendepago] TYPE_1
);

ALTER TABLE [OrdenDeCompra] ADD PRIMARY KEY ([codordendecompra]);


