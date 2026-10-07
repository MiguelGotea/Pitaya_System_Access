-- ==========================================================
-- Tabla    : EstadoInicial
-- Campos   : 8
-- Exportado: 2026-10-07 07:17:26
-- ==========================================================

CREATE TABLE [EstadoInicial] (
    [CodCajaInicial] TYPE_4,
    [Dinero] TYPE_7,
    [Fecha] TYPE_8,
    [Selladora] TYPE_4,
    [TipoCambio$_C$] TYPE_7,
    [Feriado] TYPE_1,
    [Observaciones] VARCHAR(50),
    [Eventos] VARCHAR(50)
);

ALTER TABLE [EstadoInicial] ADD PRIMARY KEY ([CodCajaInicial]);


