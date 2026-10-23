-- ==========================================================
-- Tabla    : Depositos
-- Campos   : 8
-- Exportado: 2026-10-22 22:23:19
-- ==========================================================

CREATE TABLE [Depositos] (
    [CodDeposito] TYPE_4,
    [Monto] TYPE_4,
    [Denominacion] VARCHAR(255),
    [Tipo] VARCHAR(255),
    [Fecha] TYPE_8,
    [Observacion] VARCHAR(255),
    [DuranteTurno] TYPE_1,
    [Hora] TYPE_8
);

ALTER TABLE [Depositos] ADD PRIMARY KEY ([CodDeposito]);


