' ==========================================================
' Modulo  : Form_DefinirDatosCierre
' Tipo    : 100  |  Lineas: 34
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database


Private Sub botonbloquear_Click()
On Error GoTo Nulo

'Nuevo preingreso
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO CierreDiario(HoraInicial, HoraFinal, Fecha, CodOperario)" & _
" values (#" & UltimoCierre() & "#, #" & Time() & "#, #" & Date & "#," & Me.codigologin & ")"
DoCmd.SetWarnings True

Dim cierrel As Long
cierrel = CierreFinal(Date)

Exit Sub
Nulo:
MsgBox "No se ha creado registro, vuelva a intentarlo"
End Sub

Private Sub Comando1281_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery


End Sub



