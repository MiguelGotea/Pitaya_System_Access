-- ==========================================================
-- Tabla    : ChatPitaya
-- Campos   : 4
-- Exportado: 2026-10-14 21:56:14
-- ==========================================================

CREATE TABLE [ChatPitaya] (
    [CodMensaje] TYPE_4,
    [Mensaje] VARCHAR(50),
    [Receptor] TYPE_4,
    [Fecha] TYPE_8
);

ALTER TABLE [ChatPitaya] ADD PRIMARY KEY ([CodMensaje]);


