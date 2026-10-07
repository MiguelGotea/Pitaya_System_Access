' ==========================================================
' Modulo  : Form_ResumenPreIngresosPitaya
' Tipo    : 100
' Lineas  : 21
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database


Private Sub Comando216_Click()

If IsNull(Me.adestino) Then
    Me.Filter = "[sema]=" & Me.asemana
    Me.FilterOn = True
    Me.Requery
    
Else
    Me.Filter = "[Destino] = '" & Me.adestino & "' AND [sema]=" & Me.asemana
    Me.FilterOn = True
    Me.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub
