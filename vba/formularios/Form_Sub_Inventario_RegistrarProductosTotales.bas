' ==========================================================
' Modulo  : Form_Sub_Inventario_RegistrarProductosTotales
' Tipo    : 100  |  Lineas: 58
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()

If IsNull(Me.Unidadx) Or Me.Unidadx = "" Or Me.Unidadx = " " Then
    MsgBox "no se puede editar este dato ya que no se cuenta en esta presentacion, ingresar en la otra presentacion"
    Exit Sub
End If

On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombrex) & " " & UCase(Me.Unidadx), "Cantidad", 0)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, cantidadunidad, cantidadpaquete, Fecha, CodOperario, lista)" & _
" values (" & Me.CodCotizacion & "," & tota - Me.SumaDecantidadunidad & ", 0, #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", " & Me.listaoculta & ")"
DoCmd.SetWarnings True

Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub



Private Sub Comando354_Click()

If IsNull(Me.UnidadPaquetex) Or Me.UnidadPaquetex = "" Or Me.UnidadPaquetex = " " Then
    MsgBox "no se puede editar este dato ya que no se cuenta en esta presentacion, ingresar en la otra presentacion"
    Exit Sub
End If
On Error GoTo AgainAgain
Dim tota As Double
'Me.SumaDecantidpaquete
'Me.CodCotizacion
tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombrex) & " " & UCase(Me.UnidadPaquetex), "Cantidad", 0)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, cantidadpaquete, cantidadunidad, Fecha, CodOperario, lista)" & _
" values (" & Me.CodCotizacion & ", " & tota - Me.SumaDecantidadpaquete & ", 0, #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", " & Me.listaoculta & ")"
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
