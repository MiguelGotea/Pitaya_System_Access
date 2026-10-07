' ==========================================================
' Modulo  : Form_RegistroRetiroEfectivoDolares
' Tipo    : 100  |  Lineas: 55
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database
Private Sub d1_Exit(Cancel As Integer)
If IsNull(Me.d1) Then
    Me.d1 = 0
End If
End Sub

Private Sub d5_Exit(Cancel As Integer)
If IsNull(Me.d5) Then
    Me.d5 = 0
End If
End Sub

Private Sub d10_Exit(Cancel As Integer)
If IsNull(Me.d10) Then
    Me.d10 = 0
End If
End Sub

Private Sub d20_Exit(Cancel As Integer)
If IsNull(Me.d20) Then
    Me.d20 = 0
End If
End Sub

Private Sub d50_Exit(Cancel As Integer)
If IsNull(Me.d50) Then
    Me.d50 = 0
End If
End Sub

Private Sub d100_Exit(Cancel As Integer)
If IsNull(Me.d100) Then
    Me.d100 = 0
End If
End Sub

Private Sub Comando50_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO Depositos(Monto, Denominacion, Tipo, Fecha, Hora, Observacion, DuranteTurno) values" & _
" (" & Me.Monto & ", 'Dolares', 'Deposito', #" & Me.Fecha & "#, #" & Me.horaimpresion & "#,'" & Me.Observacion & "', -1)"
DoCmd.SetWarnings True

DoCmd.PrintOut
MsgBox "Deposito Guardado Correctamente"
'Call EnviarMensajeTelegram("Aligeramiento de Efectivo en Dolares: " & Me.Monto, grupotgerencia())
DoCmd.Close acForm, "RegistroRetiroEfectivoDolares"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
