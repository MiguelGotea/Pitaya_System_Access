-- ==========================================================
-- Tabla    : ChatPitaya
-- Campos   : 4
-- Exportado: 2026-10-07 07:17:25
-- ==========================================================

CREATE TABLE [ChatPitaya] (
    [CodMensaje] TYPE_4,
    [Mensaje] VARCHAR(50),
    [Receptor] TYPE_4,
    [Fecha] TYPE_8
);

ALTER TABLE [ChatPitaya] ADD PRIMARY KEY ([CodMensaje]);


