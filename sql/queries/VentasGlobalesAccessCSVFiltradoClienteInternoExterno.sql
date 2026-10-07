-- ==========================================================
-- Consulta : VentasGlobalesAccessCSVFiltradoClienteInternoExterno
-- Exportado: 2026-10-07 06:21:34
-- ==========================================================
SELECT v1.*, (
        SELECT SUM(IIf(v2.Anulado=0, v2.PuntosLinea, 0))
        FROM VentasGlobalesFiltradoClienteInternoExterno AS v2
        WHERE (v2.Fecha + v2.Hora) <= (v1.Fecha + v1.Hora)
          AND v2.CodCliente = v1.CodCliente
    ) AS PuntosAcumulados
FROM VentasGlobalesFiltradoClienteInternoExterno AS v1
ORDER BY v1.Fecha, v1.Hora;

