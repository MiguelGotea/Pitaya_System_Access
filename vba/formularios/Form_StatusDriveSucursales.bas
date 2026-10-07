' ==========================================================
' Modulo  : Form_StatusDriveSucursales
' Tipo    : 100  |  Lineas: 14
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database


Private Sub Comando544_Click()
Call EnviarMensajeTelegram(Me.BotTelegram & Chr(32) & "reiniciardrive ", grupotgerencia())

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

Me.Requery
End Sub

