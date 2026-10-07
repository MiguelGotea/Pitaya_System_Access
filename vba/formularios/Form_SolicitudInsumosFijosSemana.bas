' ==========================================================
' Modulo  : Form_SolicitudInsumosFijosSemana
' Tipo    : 100  |  Lineas: 29
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad de productos a solicitar", "Cantidad", 0)

DoCmd.SetWarnings False

DoCmd.RunSQL "INSERT INTO PedidoInsumosFijosSemana(CodCotizacion, Cantidad, Semana, Registro, CodOperario)" & _
" values (" & Me.CodCotizacion & ", " & tota & ", " & Me.semanapedido & ",#" & Now & "#, " & Me.codoperariopedido & ")"

DoCmd.SetWarnings True

Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
