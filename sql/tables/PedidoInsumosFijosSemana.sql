-- ==========================================================
-- Tabla    : PedidoInsumosFijosSemana
-- Campos   : 6
-- Exportado: 2026-10-14 21:56:16
-- ==========================================================

CREATE TABLE [PedidoInsumosFijosSemana] (
    [CodPedidoInsumosFijosSemana] TYPE_4,
    [CodCotizacion] TYPE_4,
    [Cantidad] TYPE_7,
    [Semana] TYPE_4 NOT NULL,
    [Registro] TYPE_8,
    [CodOperario] TYPE_4
);

ALTER TABLE [PedidoInsumosFijosSemana] ADD PRIMARY KEY ([CodPedidoInsumosFijosSemana]);


