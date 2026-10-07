' ==========================================================
' Modulo  : Form_Sub_Inventario_TemporalCotizacion
' Tipo    : 100
' Lineas  : 65
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database

Private Sub Cantidad_Exit(Cancel As Integer)

On Error GoTo Nulo

Dim cantic As Double
Dim cantip As Integer
llave = Me.CodICotizacion

' Obtener la cantidad ingresada en el cuadro de texto
cantic = Me.Cantidad.Value

' Actualizar la cantidad en la tabla [Inventario Cotizacion Temporal] para el registro actual
'DoCmd.SetWarnings False
'DoCmd.RunSQL "UPDATE [Inventario Cotizacion Temporal] SET Cantidad = " & cantic & _
'    " WHERE CodCotizacion = " & Me.CodCotizacion
'DoCmd.SetWarnings True

' Comprobar si el producto es una mezcla
If ExisteMezcla(Me.CodCotizacion) = 1 Then
    ' Si es una mezcla, modificar la cantidad en la mezcla y sus subproductos
    cantip = cantidadProductosMezcla(Me.CodCotizacion)
    
    For sp = 1 To cantip
        ' Actualizar la cantidad en subproductos relacionados de la tabla [Inventario Cotizacion Temporal]
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE [Inventario Cotizacion Temporal] SET Cantidad = " & cantic & _
            " WHERE CodICotizacion = " & llave + sp
        DoCmd.SetWarnings True
    Next sp
End If

'Me.Requery
Exit Sub

Nulo:
MsgBox "Volver a intentar"

End Sub

Private Sub Comando116_Click()
On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad de productos como estandar por pedido", "Cantidad", 0)

DoCmd.SetWarnings False

DoCmd.RunSQL "INSERT INTO EstandarInsumosFijos(CodCotizacion, Cantidad, Registro)" & _
" values (" & Me.CodCotizacion & ", " & tota & ",#" & Now & "#)"

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
