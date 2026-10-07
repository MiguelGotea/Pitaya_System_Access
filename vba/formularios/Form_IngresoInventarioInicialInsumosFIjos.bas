' ==========================================================
' Modulo  : Form_IngresoInventarioInicialInsumosFIjos
' Tipo    : 100  |  Lineas: 32
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim tota As Double
Dim asctuali As Double

actuali = StockSinProcesar(Me.CodCotizacion, Me.asemana)

tota = InputBox("Ingrese la cantidad, solamente se cuentas PAQUETES SELLADOS y/o NO EN USO", "Cantidad", 0)

DoCmd.SetWarnings False

DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha, lista, CodOperario)" & _
" values (" & Me.CodCotizacion & ", " & tota - actuali & ",#" & Me.fecharealinventario & "#, 5, " & Me.codigologin & ")"

DoCmd.SetWarnings True

Me.Texto231.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
