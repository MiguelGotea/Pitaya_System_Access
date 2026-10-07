' ==========================================================
' Modulo  : Form_Ingreso Inventario Almacen
' Tipo    : 100
' Lineas  : 152
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Private Sub CrearListaUnidadesCotizacion(fech As Date)
'Lista completa ctizaciones, ingredientes Inventario SI
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT DBIngredientes.TIPO2, DBIngredientes.Tipo, Cotizaciones.Subproducto," & _
" nombreproductocoti([Cotizaciones]![CodCotizacion]) AS nombre, Cotizaciones.Inventario, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((Cotizaciones.Inventario) = True))" & _
" ORDER BY DBIngredientes.TIPO2 DESC , DBIngredientes.Tipo DESC , Cotizaciones.Subproducto," & _
" nombreproductocoti([Cotizaciones]![CodCotizacion])"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Fecha) values (" & rst("CodCotizacion") & ", #" & fech & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

miSQL = "SELECT DBIngredientes.Tipo, DBIngredientes.Nombre, DBIngredientes.Inventario, DBIngredientes.CodIngrediente" & _
" FROM DBIngredientes WHERE (((DBIngredientes.Inventario) = True))" & _
" ORDER BY DBIngredientes.Tipo, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Ingrediente](CodIngrediente, Fecha) values ('" & rst("CodIngrediente") & "', #" & fech & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

MsgBox "Se agregaron los productos de manera automatica"

Exit Sub

Nulo:
MsgBox "Error al crear datos"

End Sub


Private Sub Comando106_Click()
DoCmd.OpenForm "IngresoDePorciones", acNormal
[Forms]![IngresoDePorciones]![Procedencia] = "[Inventario Cotizacion]"
[Forms]![IngresoDePorciones]![fechaprocedencia] = Me.fechaac
End Sub

Private Sub Comando155_Click()
DoCmd.OpenForm "RegistrarProductosPorciones"
[Forms]![RegistrarProductosPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosPorciones]![adesdetabla] = "[Inventario Cotizacion]"
[Forms]![RegistrarProductosPorciones]![adesdeform] = "[Ingreso Inventario Almacen]"
End Sub

Private Sub Comando174_Click()
DoCmd.OpenForm "RegistrarProductosMostrador"
[Forms]![RegistrarProductosMostrador]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosMostrador]![adesdetabla] = "[Inventario Cotizacion]"
[Forms]![RegistrarProductosMostrador]![adesdeform] = "[Ingreso Inventario Almacen]"
End Sub

Private Sub Comando175_Click()
DoCmd.OpenForm "IngresoCodigoBarras"
[Forms]![IngresoCodigoBarras]![fechaprocedencia] = Me.fechaac
[Forms]![IngresoCodigoBarras]![adesde] = "[Inventario Cotizacion]"
End Sub



Private Sub Comando190_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado", acNormal
[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[Inventario Cotizacion]"
[Forms]![IngresoAutomaticoProductosFiltrado]![adesdeform] = "[Ingreso Inventario Almacen]"
End Sub

Private Sub Comando234_Click()
DoCmd.OpenForm "RegistrarProductosNoPorciones"
[Forms]![RegistrarProductosNoPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosNoPorciones]![adesdetabla] = "[Inventario Cotizacion]"
[Forms]![RegistrarProductosNoPorciones]![adesdeform] = "[Ingreso Inventario Almacen]"
End Sub

Private Sub Comando248_Click()

On Error GoTo Nulo

Dim cantix As Double
cantix = InputBox("Ingresar la cantidad: ", "Cantidad de productos", 0)

If cantix = 0 Then
    MsgBox "NO se agrego ningun producto"
    Exit Sub
End If

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha)" & _
" values (" & Me.codigoespecial & ", " & cantix & ", #" & Me.fechaac & "#)"
DoCmd.SetWarnings True
    
Exit Sub
Nulo:
MsgBox "Ingresar los datos correctamente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

Me.Subformulario_Inventario_Cotizacion.Requery
End Sub

Private Sub Comando21_Click()
Me.fechaac = Me.fechaac - 1
Me.Subformulario_Inventario_Cotizacion.Requery
Me.Requery
End Sub

Private Sub Comando23_Click()
Me.fechaac = Me.fechaac + 1
Me.Subformulario_Inventario_Cotizacion.Requery
Me.Requery
End Sub

Private Sub Comando40_Click()
If MsgBox("Desea generar lista completa de productos?", vbYesNo + vbQuestion) = vbYes Then
    Call CrearListaUnidadesCotizacion(Me.fechaac)
End If

Me.Subformulario_Inventario_Cotizacion.Requery
End Sub

Private Sub fechaac_Change()
Form.Refresh
End Sub

Private Sub Form_Close()
Me.Subformulario_Inventario_Cotizacion.Form.Filter = False

End Sub





