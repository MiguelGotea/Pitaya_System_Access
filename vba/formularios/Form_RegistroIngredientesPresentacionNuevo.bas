' ==========================================================
' Modulo  : Form_RegistroIngredientesPresentacionNuevo
' Tipo    : 100
' Lineas  : 104
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database



Private Sub Comando36_Click()
'Validando vacios
If IsNull(Me.codi) Then
    MsgBox "Ingresar un codigo"
    Exit Sub
End If

If metodo = 1 Then 'busca codigo solo si es emtodo1 , emtodo de agregar ingrediente + caotizacion
    If Nz(DLookup("[CodIngrediente]", "[DBIngredientes]", "[CodIngrediente]='" & Me.codi & "'"), 0) = 0 Then
        'No existe el codigo de ingrediente procede
    Else
        MsgBox "Ya existe el codigo de insumo, ingresar otro codigo"
        Me.codi = ""
        Exit Sub
    End If
End If

If IsNull(Me.nomb) Then
    MsgBox "Ingresar un nombre"
    Exit Sub
End If

If IsNull(Me.unid) Then
    MsgBox "Ingresar unidad del insumo proncipal"
    Exit Sub
End If

If IsNull(Me.tipoc) Then
    MsgBox "Ingresar el tipo de costo"
    Exit Sub
End If

If IsNull(Me.subc) Then
    MsgBox "Ingresar Sub Cuenta"
    Exit Sub
End If

If IsNull(Me.cont) Then
    MsgBox "Ingresar control de almacen"
    Exit Sub
End If

If IsNull(Me.conve) Then
    MsgBox "Ingresar el valor de conversion equivalente"
    Exit Sub
End If

If IsNull(Me.unco) Then
    MsgBox "Ingresar unidad de la presentacion a crear"
    Exit Sub
End If

'If IsNull(Me.prese) Then
'    MsgBox "Ingresar contenido detallado de presentacion"
'    Exit Sub
'End If

'If IsNull(Me.marc) Then
'    MsgBox "Ingresar Marca"
'    Exit Sub
'End If

If IsNull(Me.alma) Then
    MsgBox "Ingresar ubicacion final de almacenamiento"
    Exit Sub
End If




'Pasado todos los filtros se genera
DoCmd.SetWarnings False

If metodo = 1 Then
    DoCmd.RunSQL "INSERT INTO DBIngredientes(CodIngrediente, Nombre, NombreSinProcesar," & _
                 " Unidad, TIPO1, ComprasSemana, CodCeCoSubCuentas,CodControlAlmacenes)" & _
                 " values ('" & Me.codi & "', '" & Me.nomb & "', '" & Me.nomb & "'," & _
                 " '" & Me.unid & "', '" & Me.tipoc & "', 1, " & Me.subc & ", " & Me.cont & ")"
End If

DoCmd.RunSQL "INSERT INTO Cotizaciones(CodIngrediente, Marca, Capacidad, Conversion," & _
             " Unidad, PaquetePorciones, CodAlmacenamiento)" & _
             " values ('" & Me.codi & "','" & Me.marc & "', '" & Me.prese & "', " & Me.conve & "," & _
             " '" & Me.unco & "', 1, " & Me.alma & ")"

DoCmd.SetWarnings True

MsgBox "Producto registrado correctamente"
[Forms]![RelacionIngredientesPresentacion].Form.Sub_RelacionIngredientesPresentacion.Requery
DoCmd.Close
End Sub

Private Sub Comando37_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "Batidos Pitaya"
End Sub

