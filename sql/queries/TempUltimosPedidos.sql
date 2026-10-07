-- ==========================================================
-- Consulta : TempUltimosPedidos
-- Exportado: 2026-10-07 07:17:23
-- ==========================================================

SELECT TOP 5 SubPedido.CodSubPedido, FormatDateTime(DLookUp("[Hora]","[NotaDePedido]","[NotaDePedido]![CodPedido]=" & [SubPedido]![CodPedido]),4) & "    " & DLookUp("[Nombre]","[DBBatidos]","[DBBatidos]![CodBatido]='" & [SubPedido]![CodBatido] & "'") AS NyH
FROM SubPedido
ORDER BY SubPedido.CodSubPedido DESC;

