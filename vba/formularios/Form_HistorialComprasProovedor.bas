' ==========================================================
' Modulo  : Form_HistorialComprasProovedor
' Tipo    : 100  |  Lineas: 10
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database

Private Sub Comando2130_Click()
DoCmd.OpenForm "AutogenerarIngresodeCompras", , , "[Fecha]=#" & Me.Fecha & "# AND [CodProveedor]=" & Me.CodProveedor
[Forms]![AutogenerarIngresodeCompras]![Comando120].Visible = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
