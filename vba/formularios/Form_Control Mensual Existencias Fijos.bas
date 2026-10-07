' ==========================================================
' Modulo  : Form_Control Mensual Existencias Fijos
' Tipo    : 100
' Lineas  : 39
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub Comando52_Click()
Me.Requery
Me.Requery
End Sub

Private Sub Comando68_Click()

If mesac.Value = 1 Then
    mesac.Value = 12
    añoac.Value = añoac.Value - 1
    Me.Requery
Else
    mesac.Value = mesac.Value - 1
    Me.Requery
End If

End Sub


Private Sub Comando71_Click()

If mesac.Value = 12 Then
    mesac.Value = 1
    añoac.Value = añoac.Value + 1
    Me.Requery
Else
    mesac.Value = mesac.Value + 1
    Me.Requery
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

Me.tiempoactual = Month(Date) & " / " & Year(Date)
End Sub
