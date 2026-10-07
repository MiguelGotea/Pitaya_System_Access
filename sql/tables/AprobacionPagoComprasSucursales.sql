-- ==========================================================
-- Tabla    : AprobacionPagoComprasSucursales
-- Campos   : 5
-- Exportado: 2026-10-07 07:17:24
-- ==========================================================

CREATE TABLE [AprobacionPagoComprasSucursales] (
    [CodAprobacionPagoComprasSucursales] TYPE_4,
    [CodIngresoAlmacen] TYPE_4,
    [Sucursal] TYPE_4,
    [Aprobado] TYPE_8,
    [codordendecompra] TYPE_4
);

ALTER TABLE [AprobacionPagoComprasSucursales] ADD PRIMARY KEY ([CodAprobacionPagoComprasSucursales]);


