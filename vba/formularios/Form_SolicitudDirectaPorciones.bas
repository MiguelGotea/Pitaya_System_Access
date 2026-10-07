' ==========================================================
' Modulo  : Form_SolicitudDirectaPorciones
' Tipo    : 100
' Lineas  : 40
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database





Public Sub Comando102_Click()
Me.Requery


End Sub

Private Sub Comando116_Click()
On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad de productos como estandar por pedido", "Cantidad", 0)

DoCmd.SetWarnings False

DoCmd.RunSQL "INSERT INTO EstandarInsumosFijos(CodCotizacion, Cantidad, Registro)" & _
" values (" & Me.codporcion & ", " & tota & ",#" & Now & "#)"

DoCmd.SetWarnings True

Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub




Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
End Sub
