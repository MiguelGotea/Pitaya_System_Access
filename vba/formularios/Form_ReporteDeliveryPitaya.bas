' ==========================================================
' Modulo  : Form_ReporteDeliveryPitaya
' Tipo    : 100  |  Lineas: 27
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database















Private Sub Comando175_Click()
Me.Sub_ReporteDeliveryPitaya_Motorizado.Requery
Me.Sub_ReporteDeliveryPitaya_Sucursal.Requery
'Me.Sub_ReporteDeliveryPitaya_Sucursal.height = alturaconsultasql(Me.Sub_ReporteDeliveryPitaya_Sucursal.Form.RecordSource, 0.25)
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"


End Sub
