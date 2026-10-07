' ==========================================================
' Modulo  : Form_PruebaVelocidadFormulas
' Tipo    : 100  |  Lineas: 25
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database



Private Sub Comando105_Click()
Me.ainicio = ""
Me.afinal = ""

End Sub

Private Sub Comando92_Click()
Dim cont As Integer
cont = 1
Me.ainicio = Now

Do While cont < Me.arepe
    cont = cont + 1
    'Call Test
Loop
Me.afinal = Now
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
