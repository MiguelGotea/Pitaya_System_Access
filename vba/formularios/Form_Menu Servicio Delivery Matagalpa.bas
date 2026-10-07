' ==========================================================
' Modulo  : Form_Menu Servicio Delivery Matagalpa
' Tipo    : 100
' Lineas  : 20
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================


Private Sub Comando2313_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, Vinculo)" & _
" VALUES ('C164',1, " & [Forms]![Nota de Pedido]![CodPedido] & ", 0, '-', 0)"
DoCmd.SetWarnings True


End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

Private Sub regresar_Click()
DoCmd.Close
End Sub
