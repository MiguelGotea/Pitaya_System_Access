' ==========================================================
' Modulo  : Form_Control de Servicios Globales
' Tipo    : 100
' Lineas  : 39
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Private Sub Comando21_Click()
Me.Requery
End Sub

Private Sub Comando68_Click()

If ames.Value = 1 Then
    ames.Value = 12
    aano.Value = aano.Value - 1
    Me.Requery
Else
    ames.Value = ames.Value - 1
    Me.Requery
End If

End Sub

Private Sub Comando71_Click()

If ames.Value = 12 Then
    ames.Value = 1
    aano.Value = aano.Value + 1
    Me.Requery
Else
    ames.Value = ames.Value + 1
    Me.Requery
End If

End Sub

Private Sub Form_Load()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
