' ==========================================================
' Modulo  : Form_Sub_ReporteDeliveryPitaya_Motorizado
' Tipo    : 100  |  Lineas: 26
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database















Private Sub Comando148_Click()
DoCmd.OpenForm "HistorialVentasAtencionAlClienteFiltro", acNormal, , "[CodigoMotorizado]=" & Me.CodigoMotorizado & " AND [Fecha]>=#" & [Forms]![ReporteDeliveryPitaya]![fechadesde] & "# AND [Fecha]<=#" & [Forms]![ReporteDeliveryPitaya]![fechahasta] & "#"
'DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]<>0 AND [Transferencia]=0"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"


End Sub
