' ==========================================================
' Modulo  : Form_RegistroCeCoSubCuentas
' Tipo    : 100  |  Lineas: 21
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:08
' ==========================================================

Option Compare Database


Private Sub Comando108_Click()
If Me.acuenta = "" Or IsNull(Me.acuenta) Then
    Me.FilterOn = False
    Me.Requery
Else
    Me.Filter = "[CodCeCoCuentas]=" & Me.acuenta
    Me.FilterOn = True
End If
Me.Requery



End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "Batidos Pitaya"
Me.ShortcutMenu = False
End Sub
