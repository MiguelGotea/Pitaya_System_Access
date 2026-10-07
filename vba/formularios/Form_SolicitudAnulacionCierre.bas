' ==========================================================
' Modulo  : Form_SolicitudAnulacionCierre
' Tipo    : 100  |  Lineas: 33
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database

Private Sub botonbloquear_Click()
On Error GoTo Nulo

If IsNull(Me.pedidomotivo) Or Me.pedidomotivo = " " Then
    MsgBox "Ingresar motivo de anulacion"
    Exit Sub
End If

'Dim mensajeg As String
'mensajeg = "Solicitud de Reinicio de Cierre: " & Me.pedidoanular & vbCrLf & vbCrLf & _
          "Solicitado por: " & Me.pedidooperario & vbCrLf & _
          "Motivo: " & Me.pedidomotivo

'Call EnviarMensajeTelegram(mensajeg, grupotanulaciones())
MsgBox "Solicitud para reiniciar cierre enviado correctamente, el cierre estara disponible en breve"
DoCmd.Close acForm, "SolicitudAnulacionCierre"
Exit Sub
Nulo:
MsgBox "Solicitar reinicio de cierre directamente, problemas con el servidor"

End Sub

Private Sub Comando1253_Click()
DoCmd.Close acForm, "SolicitudAnulacionCierre"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

