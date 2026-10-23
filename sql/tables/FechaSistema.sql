-- ==========================================================
-- Tabla    : FechaSistema
-- Campos   : 4
-- Exportado: 2026-10-22 22:23:19
-- ==========================================================

CREATE TABLE [FechaSistema] (
    [Dates] TYPE_8 NOT NULL,
    [semana] TYPE_4,
    [mes] TYPE_4,
    [anio] TYPE_4
);

ALTER TABLE [FechaSistema] ADD PRIMARY KEY ([Dates]);


