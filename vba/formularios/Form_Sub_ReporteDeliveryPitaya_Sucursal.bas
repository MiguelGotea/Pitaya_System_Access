' ==========================================================
' Modulo  : Form_Sub_ReporteDeliveryPitaya_Sucursal
' Tipo    : 100
' Lineas  : 26
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database















Private Sub Comando148_Click()
DoCmd.OpenForm "HistorialVentasAtencionAlClienteFiltro", acNormal, , "[CodLocal]=" & Me.CodLocal & " AND [TipoPedido]='" & Me.TipoPedido & "' AND [Fecha]>=#" & [Forms]![ReporteDeliveryPitaya]![fechadesde] & "# AND [Fecha]<=#" & [Forms]![ReporteDeliveryPitaya]![fechahasta] & "#"

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"


End Sub
