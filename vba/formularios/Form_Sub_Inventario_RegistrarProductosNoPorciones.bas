' ==========================================================
' Modulo  : Form_Sub_Inventario_RegistrarProductosNoPorciones
' Tipo    : 100
' Lineas  : 46
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombrex) & " " & UCase(Me.Unidadx), "Cantidad", 0)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, cantidadunidad, Fecha, CodOperario, lista)" & _
" values (" & Me.CodCotizacion & "," & tota - Me.SumaDecantidadunidad & ", #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", 3)"
DoCmd.SetWarnings True

Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub



Private Sub Comando354_Click()
On Error GoTo AgainAgain
Dim tota As Double
'Me.SumaDecantidpaquete
'Me.CodCotizacion
tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombrex) & " " & UCase(Me.UnidadPaquetex), "Cantidad", 0)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, cantidadpaquete, Fecha, CodOperario, lista)" & _
" values (" & Me.CodCotizacion & ", " & tota - Me.SumaDecantidadpaquete & ", #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", 3)"
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
