' ==========================================================
' Modulo  : Form_ElegirTipoSalidaEfectivo
' Tipo    : 100  |  Lineas: 18
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:06
' ==========================================================

Option Compare Database



Private Sub Comando65_Click()
DoCmd.OpenForm "RegistroRetiroEfectivoCordobas"
[Forms]![RegistroRetiroEfectivoCordobas]![codigologin] = Me.codigologin
End Sub

Private Sub Comando66_Click()
DoCmd.OpenForm "RegistroRetiroEfectivoDolares"
[Forms]![RegistroRetiroEfectivoDolares]![codigologin] = Me.codigologin
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
