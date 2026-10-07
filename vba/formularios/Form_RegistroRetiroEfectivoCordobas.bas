' ==========================================================
' Modulo  : Form_RegistroRetiroEfectivoCordobas
' Tipo    : 100
' Lineas  : 73
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Private Sub c0_5_Exit(Cancel As Integer)
If IsNull(Me.[c0.5]) Then
    Me.[c0.5] = 0
End If
End Sub

Private Sub c1_Exit(Cancel As Integer)
If IsNull(Me.c1) Then
    Me.c1 = 0
End If
End Sub

Private Sub c5_Exit(Cancel As Integer)
If IsNull(Me.c5) Then
    Me.c5 = 0
End If
End Sub

Private Sub c10_Exit(Cancel As Integer)
If IsNull(Me.c10) Then
    Me.c10 = 0
End If
End Sub

Private Sub c20_Exit(Cancel As Integer)
If IsNull(Me.c20) Then
    Me.c20 = 0
End If
End Sub

Private Sub c50_Exit(Cancel As Integer)
If IsNull(Me.c50) Then
    Me.c50 = 0
End If
End Sub

Private Sub c100_Exit(Cancel As Integer)
If IsNull(Me.c100) Then
    Me.c100 = 0
End If
End Sub

Private Sub c200_Exit(Cancel As Integer)
If IsNull(Me.c200) Then
    Me.c200 = 0
End If
End Sub

Private Sub c500_Exit(Cancel As Integer)
If IsNull(Me.c500) Then
    Me.c500 = 0
End If
End Sub


Private Sub Comando50_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO Depositos(Monto, Denominacion, Tipo, Fecha, Hora, Observacion, DuranteTurno) values" & _
" (" & Me.Monto & ", 'Cordobas', 'Deposito', #" & Me.Fecha & "#, #" & Me.horaimpresion & "#, '" & Me.Observacion & "', -1)"
DoCmd.SetWarnings True

DoCmd.PrintOut
MsgBox "Deposito Guardado Correctamente"
'Call EnviarMensajeTelegram("Aligeramiento de Efectivo en Cordobas: " & Me.Monto, grupotgerencia())
DoCmd.Close acForm, "RegistroRetiroEfectivoCordobas"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
