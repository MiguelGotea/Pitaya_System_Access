' ==========================================================
' Modulo  : Form_Ingreso Inventario Pitaya
' Tipo    : 100
' Lineas  : 961
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Public Function cargarllaveatemporal()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal] SELECT * FROM [Inventario Cotizacion]"
DoCmd.SetWarnings True

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Ingrediente Temporal] SELECT * FROM [Inventario Ingrediente]"
DoCmd.SetWarnings True

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM [Inventario Cotizacion Temporal]"
DoCmd.SetWarnings True

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM [Inventario Ingrediente Temporal]"
DoCmd.SetWarnings True

End Function

Function TrasladarTemporalAOficial()
'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Registro As Long
Dim canti As Double
Dim cotix As Integer
Dim rsPorciones As DAO.Recordset
Dim cantip As Integer
Dim list As Integer
Dim fechi As Date

miSQL = "SELECT [Inventario Cotizacion Temporal].CodICotizacion, [Inventario Cotizacion Temporal].CodCotizacion," & _
" [Inventario Cotizacion Temporal].lista, [Inventario Cotizacion Temporal].cantidadunidad," & _
" [Inventario Cotizacion Temporal].cantidadpaquete, [Inventario Cotizacion Temporal].Fecha, [Inventario Cotizacion Temporal].CodOperario" & _
" FROM [Inventario Cotizacion Temporal]" & _
" ORDER BY [Inventario Cotizacion Temporal].CodICotizacion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    Registro = rst("CodICotizacion")
    cotix = rst("CodCotizacion")
    list = rst("Lista")
    fechi = rst("Fecha")
    oper = rst("CodOperario")
    canti = rst("cantidadunidad") * factorinventariounidadunitario(cotix) + rst("cantidadpaquete") * factorinventariounidaddespacho(cotix)

    If canti <> 0 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE [Inventario Cotizacion Temporal] SET Cantidad = " & canti & _
            " WHERE CodICotizacion = " & Registro
        DoCmd.SetWarnings True
        
        If ExisteMezcla(cotix) = 1 Then
            ' Agregar la mezcla original
            ' Obtener CodCotizacionPorcion relacionados

            Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & cotix)
    
            ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
            Do While Not rsPorciones.EOF

                If IsNull(oper) Then
                    oper = [Forms]![Ingreso Inventario Pitaya]![codigologin]
                End If
                DoCmd.SetWarnings False
                DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, cantidadunidad, Cantidad, Fecha, lista, CodOperario)" & _
                    " values (" & rsPorciones!CodCotizacionPorcion & ", " & canti & ", " & canti & ", #" & fechi & "#, " & list & ", " & oper & ")"
                DoCmd.SetWarnings True
                rsPorciones.MoveNext
            Loop
    
            rsPorciones.Close
        
        End If
    
    End If
    
    rst.MoveNext
    
Loop

rst.Close
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion] (CodCotizacion, Cantidad, Fecha, lista, CodOperario, primerenvio, segundoenvio, cantidadunidad, cantidadpaquete) " & _
"SELECT CodCotizacion, Cantidad, Fecha, lista, CodOperario, primerenvio, segundoenvio, cantidadunidad, cantidadpaquete FROM [Inventario Cotizacion Temporal]"
DoCmd.SetWarnings True

Exit Function

Nulo:
MsgBox "Problemas al mover inventario termporal a final"

End Function

Function TrasladarIngredienteTemporalAOficial()

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Ingrediente] SELECT * FROM [Inventario Ingrediente Temporal]"
DoCmd.SetWarnings True

End Function


Function EliminarInventario()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM [Inventario Cotizacion Temporal]"
DoCmd.SetWarnings True



End Function

Public Sub bloquearfecha()
Me.fechaac.Enabled = False
Me.Comando21.Enabled = False
Me.Comando23.Enabled = False
End Sub

Private Sub desbloquearfecha()
Me.fechaac.Enabled = True
Me.Comando21.Enabled = True
Me.Comando23.Enabled = True
End Sub

Private Sub CrearListaNoPorciones(fech As Date)
'Lista completa ctizaciones, ingredientes Inventario SI
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control, [DBBatidos]![CodGrupo]=7 AS nomostrador," & _
" DBBatidos.Vigencia, Cotizaciones.Subproducto, SubReceta.codporcion, [Cotizaciones]![Marca] & ' '='Almacen Global ' AS NoAlmacenGlobal," & _
" Cotizaciones.Prioridad, Cotizaciones.Descontinuado, Cotizaciones.CodCotizacion" & _
" FROM ((((DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo)" & _
" INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control," & _
" [DBBatidos]![CodGrupo]=7, DBBatidos.Vigencia, Cotizaciones.Subproducto, SubReceta.codporcion," & _
" [Cotizaciones]![Marca] & ' '='Almacen Global ', Cotizaciones.Prioridad, Cotizaciones.Descontinuado, Cotizaciones.CodCotizacion" & _
" HAVING (((Grupos.Control) = True) And ((TiposVariables.Control) = True)" & _
" And (([DBBatidos]![CodGrupo] = 7) = False) And ((DBBatidos.Vigencia) = True)" & _
" And ((Cotizaciones.Subproducto) = False) And ((SubReceta.codporcion) Is Null)" & _
" And (([Cotizaciones]![Marca] & ' ' = 'Almacen Global ') = False) And ((Cotizaciones.Prioridad) = True)" & _
" And ((Cotizaciones.Descontinuado) = False))" & _
" ORDER BY TiposVariables.Orden, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
contador = 0

Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Fecha, lista, CodOperario)" & _
    " values (" & rst("CodCotizacion") & ", #" & fech & "#, 3, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ")"
    DoCmd.SetWarnings True
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

'MsgBox "Se agregaron los productos de manera automatica"

Exit Sub

Nulo:
MsgBox "Error al crear datos"

End Sub

Private Sub CrearListaMostrador(fech As Date)
'Lista completa ctizaciones, ingredientes Inventario SI
On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

Dim cantli As Integer
Dim cot As Integer
Dim canti As Double

miSQL = "SELECT DBBatidos.Marca, DBBatidos.CodSubGrupo, DBBatidos.CodGrupo, DBBatidos.Vigencia," & _
" SubReceta.InsumoClave, SubReceta.codporcion, SubReceta.CodIngrediente" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY DBBatidos.Marca, DBBatidos.CodSubGrupo, DBBatidos.CodGrupo, DBBatidos.Vigencia," & _
" SubReceta.InsumoClave, SubReceta.codporcion, SubReceta.CodIngrediente" & _
" HAVING (((DBBatidos.CodGrupo) = 7) And ((DBBatidos.Vigencia) = True) And ((SubReceta.InsumoClave) = True))" & _
" ORDER BY DBBatidos.Marca, DBBatidos.CodSubGrupo"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
contador = 0
    
For f = 1 To cantli

    If IsNull(rst("codporcion")) Then 'tiene codigo cotizacion directo a ingresar
        cot = CotiDirectoDeIngrediente(rst("CodIngrediente"))
    Else 'tiene codigo de porcion entonces se agrega la porcion
        cot = rst("codporcion")
    End If

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Cantidad, Fecha, lista, CodOperario)" & _
    " values (" & cot & ", 0, #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, 4, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ")"
    DoCmd.SetWarnings True
    
    ''Verificar si existe relacion con CodCotizacionPorcion
    '    If ExisteMezcla(cot) = 1 Then
    '        'Agregar la mezcla original
    '        'Obtener CodCotizacionPorcion relacionados
    '        Dim rsPorciones As dao.Recordset
    '        Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & cot)
    '
    '        'Agregar CodCotizacionPorcion relacionados con la misma cantidad
    '        Do While Not rsPorciones.EOF
    '            DoCmd.SetWarnings False
    '            DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Cantidad, Fecha, lista)" & _
    '            " values (" & rsPorciones!CodCotizacionPorcion & ", 0, #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, 4)"
    '            DoCmd.SetWarnings True
    '            rsPorciones.MoveNext
    '        Loop
    '        rsPorciones.Close
    '    End If
    rst.MoveNext
Next f

rst.Close

'MsgBox "Se agregaron los productos de manera automatica"

Exit Sub

Nulo:
MsgBox "Error al crear datos de mostrador"

End Sub

Private Sub CrearListaUnidadesGranel(fech As Date)
'Lista completa unidades a granel, 
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT DBIngredientes.Tipo, DBIngredientes.Nombre, DBIngredientes.Inventario, DBIngredientes.CodIngrediente" & _
" FROM DBIngredientes WHERE (((DBIngredientes.Inventario) = True))" & _
" ORDER BY DBIngredientes.Tipo, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
contador = 0

Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Ingrediente Temporal](CodIngrediente, Fecha) values ('" & rst("CodIngrediente") & "', #" & fech & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

'MsgBox "Se agregaron los productos de manera automatica"

Exit Sub

Nulo:
MsgBox "Error al crear datos"

End Sub

Private Sub CrearListaInsumosFijos(fech As Date)
'Lista completa unidades a consumible verdadero fijos, 
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim codtemporal As Integer

miSQL = "SELECT DBIngredientes.Consumible, Cotizaciones.Prioridad," & _
" Cotizaciones.Descontinuado, DBIngredientes.Vigente, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (((DBIngredientes.Consumible)<>0) AND ((Cotizaciones.Prioridad)<>0)" & _
" AND ((Cotizaciones.Descontinuado)=0) AND ((DBIngredientes.Vigente)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

Do While Not rst.EOF
    codtemporal = rst("CodCotizacion")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Fecha, lista, CodOperario)" & _
    " values (" & codtemporal & ", #" & fech & "#, 5, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ")"
    DoCmd.SetWarnings True
    rst.MoveNext
Loop
rst.Close

'MsgBox "Se agregaron los productos de manera automatica"

Exit Sub

Nulo:
MsgBox "Error al crear datos Fijos"

End Sub
Private Sub CrearListaPorciones(fech As Date)

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer
Dim codtemporal As Integer

miSQL = "SELECT TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion]) AS nombrefinal, TiposVariables.Control," & _
" DBBatidos.Vigencia, SubReceta.codporcion, [DBBatidos]![CodGrupo]=7 AS nomostrador," & _
" PorcionDentroDeMezcla([SubReceta]![codporcion]) AS mezcla, DBIngredientes.Tipo, DBIngredientes.Nombre" & _
" FROM ((SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" GROUP BY TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion]), TiposVariables.Control," & _
" DBBatidos.Vigencia, SubReceta.codporcion, [DBBatidos]![CodGrupo]=7," & _
" PorcionDentroDeMezcla([SubReceta]![codporcion]), DBIngredientes.Tipo, DBIngredientes.Nombre" & _
" HAVING (((TiposVariables.Control) = True) And ((DBBatidos.Vigencia) = True)" & _
" And ((SubReceta.codporcion) Is Not Null) And (([DBBatidos]![CodGrupo] = 7) = False)" & _
" And ((PorcionDentroDeMezcla([SubReceta]![codporcion])) = 0))" & _
" ORDER BY TiposVariables.Orden, nombreproductocotiprocesado([SubReceta]![codporcion])"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
contador = 0

Do While Not rst.EOF
    
    codtemporal = rst("codporcion")
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Fecha, lista, CodOperario)" & _
    " values (" & codtemporal & ", #" & fech & "#, 2, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ")"
    DoCmd.SetWarnings True
    
    'If ExisteMezcla(codtemporal) = 1 Then
    '
    '    ' Obtener CodCotizacionPorcion relacionados
    '    Dim rsPorciones As dao.Recordset
    '    Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & codtemporal)
    '
    '    ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
    '    Do While Not rsPorciones.EOF
    '
    '        DoCmd.SetWarnings False
    '        DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Fecha, lista)" & _
    '            " values (" & rsPorciones!CodCotizacionPorcion & ", #" & fech & "#, 2)"
    '        DoCmd.SetWarnings False
    '
    '        rsPorciones.MoveNext
    '    Loop
    '
    '    rsPorciones.Close
    'End If
    
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close
'MsgBox "Se agregaron " & contador & " elementos"

Exit Sub

Nulo:
MsgBox "Error al crear datos"

End Sub






Private Sub Comando190_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado", acNormal
[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[Inventario Cotizacion]"
End Sub

Public Sub Comando317_Click()

DoCmd.OpenForm "LogueoAutorizacion"
'[Forms]![LogueoUsuario]![CodigoBusqueda] = Me.codigologin
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(5, codigoLocal(), Date)
[Forms]![LogueoAutorizacion]![origenlogueo] = "DesbloquearInventario"

End Sub


Private Sub Comando330_Click()
'If Weekday(Me.fechaac) >= 2 And Weekday(Me.fechaac) <= 6 Then
'    MsgBox "No se permite registrar inventarios en dias lunes a viernes"
'    Exit Sub
'End If

Me.Estado.Value = "C"
If Me.Comando330.Caption = "NUEVO INVENTARIO" Then
    Call cargarllaveatemporal
    
    Call CrearListaPorciones(Me.fechaac)
    Call CrearListaNoPorciones(Me.fechaac)
    Call CrearListaMostrador(Me.fechaac)
    Call CrearListaInsumosFijos(Me.fechaac)
    MsgBox "Inventario Nuevo ha sido creado"
Else
    Me.fechaac = ComprobarInventario()
    Call bloquearfecha
    
End If

Me.fechaac.Enabled = False
Me.Comando21.Enabled = False
Me.Comando23.Enabled = False

Me.Comando330.Caption = "NUEVO INVENTARIO"
Me.Comando330.Enabled = False
Me.Comando366.Enabled = True
Me.Comando317.Enabled = False
Me.Comando363.Enabled = False

Call ocultartodoventanainventario
End Sub



Private Sub Comando363_Click()
If MsgBox("Desea guardar los cambios en el inventario?", vbYesNo + vbQuestion) = vbYes Then
    Call TrasladarTemporalAOficial

    Call EliminarInventario
    
    Me.Estado.Value = "A"
    Call ocultartodoventanainventario
    Call desbloquearfecha
    
    Me.Comando330.Enabled = False
    Me.Comando366.Enabled = False
    Me.Comando317.Enabled = True
    Me.Comando363.Enabled = False
    
    If EsSistemaDeTienda() Then
        Call SyncKardexInventarioCotizacion30Dias
    End If
End If


End Sub


Public Sub Comando366_Click()

If MsgBox("¿Desea guardar el Inventario?", vbQuestion + vbYesNo, "Confirmar Guardar el Inventario") = vbYes Then
    Call TrasladarTemporalAOficial
    Call desbloquearfecha
    
    Me.Estado.Value = "A"
    Me.Comando330.Enabled = False
    Me.Comando366.Enabled = False
    Me.Comando317.Enabled = True
    Me.Comando363.Enabled = False
    
    Call ocultartodoventanainventario
    Call EliminarInventario
    
    If EsSistemaDeTienda() Then
        Call SyncKardexInventarioCotizacion30Dias
    End If
End If

End Sub



Private Sub Comando452_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "DesbloquearInventarioMain"
End Sub

Private Sub fechaac_Exit(Cancel As Integer)

Call estadobotonesnuevoeditar
Call ocultartodoventanainventario

Me.Requery
End Sub

Private Sub FIjos_Click()
Call estadoBotones("Fijos")
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Estado.Value = "A"

If ComprobarInventario() = #1/1/2000# Then
    'No hay inventario temporal
    
    'Si EIO = 1
    If ExisteInventarioOficial(Date) = 1 Then
        Me.Comando330.Enabled = False
        Me.Comando366.Enabled = False
        Me.Comando317.Enabled = True
        Me.Comando363.Enabled = False
        Me.Estado.Value = "A"
    Else
    'Si EIO = 0
        Me.Comando330.Enabled = True
        Me.Comando366.Enabled = False
        Me.Comando317.Enabled = False
        Me.Comando363.Enabled = False
        Me.Estado.Value = "A"
    End If
Else
    'Hay inventario temporal
    If MsgBox("Existe un inventario en proceso de registro, ¿Desea abrirlo?", vbYesNo) = vbYes Then
        'Seguir editando inventario temporal
        Me.fechaac = ComprobarInventario()
        Call bloquearfecha
        Me.Comando330.Enabled = False
        Me.Comando366.Enabled = True
        Me.Comando317.Enabled = False
        Me.Comando363.Enabled = False
        Me.Estado.Value = "C"
    Else
    'No seguir editando inventario temporal
        'Si EIO = 1
        If ExisteInventarioOficial(Date) = 1 Then
            Me.Comando330.Enabled = False
            Me.Comando366.Enabled = False
            Me.Comando317.Enabled = True
            Me.Comando363.Enabled = False
            Me.Estado.Value = "A"
        'Si EIO = 0
        Else
            Me.Comando330.Enabled = True
            Me.Comando330.Caption = "SEGUIR CON INVENTARIO"
            Me.Comando366.Enabled = False
            Me.Comando317.Enabled = False
            Me.Comando363.Enabled = False
            Me.Estado.Value = "A"

        End If
    End If
End If

Call ocultartodoventanainventario

End Sub

Function ComprobarInventario() As Date

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [Inventario Cotizacion Temporal].CodICotizacion, 1 AS Contador, [Inventario Cotizacion Temporal].Fecha" & _
" FROM [Inventario Cotizacion Temporal]"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprobarInventario = rst("Fecha")
rst.Close

Exit Function

Nulo:
ComprobarInventario = #1/1/2000#

End Function

Function ExisteInventarioOficial(fechaIn As Date) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [Inventario Cotizacion].CodICotizacion, numerosemana([Inventario Cotizacion]![Fecha]) AS sema, 1 AS Contador" & _
" FROM [Inventario Cotizacion]" & _
" WHERE (((numerosemana([Inventario Cotizacion]![Fecha]))=numerosemana(#" & fechaIn & "#)))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ExisteInventarioOficial = rst("Contador")
rst.Close

Exit Function

Nulo:
ExisteInventarioOficial = 0

End Function



Private Sub Comando21_Click()
Me.fechaac = Me.fechaac - 1

Call estadobotonesnuevoeditar
Call ocultartodoventanainventario

Me.Requery
End Sub

Private Sub Comando23_Click()
If fechaac.Value + 1 > Date + 1 Then
    MsgBox "No se puede ver ingreso de inventario de fechas posteriores"
Else
    fechaac.Value = fechaac.Value + 1


    Call estadobotonesnuevoeditar
    Call ocultartodoventanainventario
    
    Me.Requery
End If
End Sub

Private Sub Comando40_Click()
If MsgBox("Desea generar lista completa de productos?", vbYesNo + vbQuestion) = vbYes Then
    'Call CrearListaUnidadesCotizacion(Me.fechaac)
End If

'Me.Subformulario_Inventario_Cotizacion.Requery
'Me.Subformulario_Inventario_Ingrediente.Requery
End Sub



Private Sub Form_Close()
If Me.Estado.Value = "B" Then
    If ComprobarInventario() = #1/1/2000# Then
        'Nada
    Else
        If MsgBox("¿Desea guardar los cambios realizados en el inventario?" & vbCrLf & vbCrLf & "Si: Guardar cambios en el inventario" & vbCrLf & "No: Borrar cambios realizados", vbYesNo) = vbYes Then
            Call TrasladarTemporalAOficial

            Call EliminarInventario
        Else
            Call EliminarInventario
        End If
    End If
End If
End Sub

Private Sub Ingrediente_Click()
Call estadoBotones("Ingrediente")
End Sub

Private Sub Mostrador_Click()
Call estadoBotones("Mostrador")
End Sub

Private Sub NoPorciones_Click()
Call estadoBotones("NoPorciones")
End Sub

Private Sub Porciones_Click()
Call estadoBotones("Porciones")
End Sub
Public Sub ocultartodoventanainventario()
'Me.Sub_Inventario_Ingrediente.Visible = False
'Me.Sub_Inventario_RegistrarProductosMostrador.Visible = False
'Me.Sub_Inventario_RegistrarProductosNoPorciones.Visible = False
'Me.Sub_Inventario_RegistrarProductosPorciones.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Visible = False
Me.Sub_Inventario_TemporalCotizacion.Visible = False
'Me.Sub_Inventario_Ingrediente_Temporal.Visible = False
'Me.Sub_Inventario_Ingrediente_Temporal_Nuevo.Visible = False
End Sub

Sub modovistagenerico()

Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDecantidadpaquete.Enabled = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDecantidadunidad.Enabled = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Unidadx.Enabled = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.UnidadPaquetex.Enabled = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Nombrex.Enabled = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDeprimerenvio.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDesegundoenvio.Enabled = False

Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery

Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.width = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.TopPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.BottomPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.LeftPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.RightPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.width = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.TopPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.BottomPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.LeftPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.RightPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto233.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto233.width = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto360.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto360.width = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta232.Visible = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta363.Visible = False
End Sub

Sub modoediciongenerico()
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDecantidadpaquete.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDecantidadunidad.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Unidadx.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.UnidadPaquetex.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.Nombrex.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDeprimerenvio.Enabled = False
Me.Sub_Inventario_RegistrarProductosTotales.Form.SumaDesegundoenvio.Enabled = False

Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery

Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.Visible = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.width = 1296
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.TopPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.BottomPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.LeftPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando131.RightPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.Visible = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.width = 1296
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.TopPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.BottomPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.LeftPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Comando354.RightPadding = 0
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto233.Visible = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto233.width = 1296
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto360.Visible = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Texto360.width = 1296
Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta232.Visible = True
Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta363.Visible = True
End Sub

Public Sub estadoBotones(boton As String)
Select Case boton
    Case "Porciones"
        ' Porciones según el estado actual
        Select Case Me.Estado
        Case "A"
            ' Cumple la condición según el estado A para NoPorciones
            Call ocultartodoventanainventario
            Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=2 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
            
            Call modovistagenerico
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "1. PORCIONES"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 2
            
        Case "B"
            ' Cumple la condición según el estado B para NoPorciones
            Call ocultartodoventanainventario
            Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=2 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
            
            Call modoediciongenerico
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "1. PORCIONES"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 2
            
        Case "C"
            ' Cumple la condición según el estado C para Porciones
            Call ocultartodoventanainventario
            
            Me.Sub_Inventario_TemporalCotizacion.Visible = True
            Me.Sub_Inventario_TemporalCotizacion.Form.Filter = ""
            Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = False
            
            'Boton filtrar
            'Me.Comando426.Visible = False
            
            Me.Sub_Inventario_TemporalCotizacion.Form.Filter = "[lista]=2 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([Inventario Cotizacion Temporal]![CodCotizacion]))=0)"
            Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = True
            Me.Sub_Inventario_TemporalCotizacion.Form.Requery
            'Me.Sub_Inventario_TemporalCotizacion.Form.titulotemporal = "NO PORCIONES"
        End Select

    Case "NoPorciones"
        ' Porciones según el estado actual
        Select Case Me.Estado
        Case "A"
            ' Cumple la condición según el estado A para NoPorciones
            Call ocultartodoventanainventario
            Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=3 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
            
            Call modovistagenerico
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "2. NO PORCIONES"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 3
            
        Case "B"
            
            ' Cumple la condición según el estado B para NoPorciones
            Call ocultartodoventanainventario
            Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=3 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True

            Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
            
            Call modoediciongenerico
            Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "2. NO PORCIONES"
            Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 3
            
        Case "C"
            ' Cumple la condición según el estado C para NoPorciones
            Call ocultartodoventanainventario
            
            Me.Sub_Inventario_TemporalCotizacion.Visible = True
            Me.Sub_Inventario_TemporalCotizacion.Form.Filter = ""
            Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = False
                
            Me.Sub_Inventario_TemporalCotizacion.Form.Filter = "[lista]=3 AND [Fecha]=#" & Me.fechaac & "#"
            Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = True
            Me.Sub_Inventario_TemporalCotizacion.Form.Requery
            'Me.Sub_Inventario_TemporalCotizacion.Form.titulotemporal = "PORCIONES"
        End Select

    Case "Mostrador"
        ' Mostrador según el estado actual
        Select Case Me.Estado
            Case "A"
                ' Cumple la condición según el estado A para NoPorciones
                Call ocultartodoventanainventario
                Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=4 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
                
                Call modovistagenerico
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "3. MOSTRADOR"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 4
            
            Case "B"
                ' Cumple la condición según el estado B para NoPorciones
                Call ocultartodoventanainventario
                Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=4 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
                
                Call modoediciongenerico
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "3. MOSTRADOR"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 4
                
            Case "C"
                ' Cumple la condición según el estado C para Mostrador
                Call ocultartodoventanainventario
                
                Me.Sub_Inventario_TemporalCotizacion.Visible = True
                Me.Sub_Inventario_TemporalCotizacion.Form.Filter = ""
                Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = False
                
                Me.Sub_Inventario_TemporalCotizacion.Form.Filter = "[lista]=4 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([Inventario Cotizacion Temporal]![CodCotizacion]))=0)"
                Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = True
                Me.Sub_Inventario_TemporalCotizacion.Form.Requery
                'Me.Sub_Inventario_TemporalCotizacion.Form.titulotemporal = "MOSTRADOR"
        End Select
        
    Case "Fijos"
        ' Fijos según el estado actual
        Select Case Me.Estado
            Case "A"
                ' Cumple la condición según el estado A para Fijos
                Call ocultartodoventanainventario
                Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=5 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
                
                Call modovistagenerico
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "4. FIJOS"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 5
            
            Case "B"
                ' Cumple la condición según el estado B para NoPorciones
                Call ocultartodoventanainventario
                Me.Sub_Inventario_RegistrarProductosTotales.Visible = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Filter = "[lista]=5 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([CodCotizacion]))=0)"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.FilterOn = True
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Requery
                
                Call modoediciongenerico
                Me.Sub_Inventario_RegistrarProductosTotales.Form.Etiqueta67 = "4. FIJOS"
                Me.Sub_Inventario_RegistrarProductosTotales.Form.listaoculta = 5
                
            Case "C"
                ' Cumple la condición según el estado C para Fijos
                Call ocultartodoventanainventario
                
                Me.Sub_Inventario_TemporalCotizacion.Visible = True
                Me.Sub_Inventario_TemporalCotizacion.Form.Filter = ""
                Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = False
                
                Me.Sub_Inventario_TemporalCotizacion.Form.Filter = "[lista]=5 AND [Fecha]=#" & Me.fechaac & "# AND ((PorcionDentroDeMezcla([Inventario Cotizacion Temporal]![CodCotizacion]))=0)"
                Me.Sub_Inventario_TemporalCotizacion.Form.FilterOn = True
                Me.Sub_Inventario_TemporalCotizacion.Form.Requery
                'Me.Sub_Inventario_TemporalCotizacion.Form.titulotemporal = "FIJOS"
        End Select

End Select
End Sub

Private Sub estadobotonesnuevoeditar()
' Solo aplica modo vista porque en nuevo y en eitar ya no se puede explorar en fechas

Select Case numerosemana(Me.fechaac.Value)
    Case Is > numerosemana(Date)
        'cuando se explora en dias despues de la semana actual
        Me.Comando330.Enabled = False
        Me.Comando366.Enabled = False
        Me.Comando317.Enabled = False
        Me.Comando363.Enabled = False
    Case numerosemana(Date)
        'cuando se explora en dias de la semana actual
        If ExisteInventarioOficial(Me.fechaac.Value) = 1 Then
            Me.Comando330.Enabled = False
            Me.Comando366.Enabled = False
            Me.Comando317.Enabled = True
            Me.Comando363.Enabled = False
        Else
            Me.Comando330.Enabled = True
            Me.Comando366.Enabled = False
            Me.Comando317.Enabled = False
            Me.Comando363.Enabled = False
        End If
        
    Case numerosemana(Date) - 1
        'Cuando se ecplora en dias de semana pasadas
        If Weekday(Date, 2) > 2 Then  ' ya no de puede deitar despues de martes
            Me.Comando330.Enabled = False
            Me.Comando366.Enabled = False
            Me.Comando317.Enabled = False
            Me.Comando363.Enabled = False
        Else
            If ExisteInventarioOficial(Me.fechaac.Value) = 1 Then
                Me.Comando330.Enabled = False
                Me.Comando366.Enabled = False
                Me.Comando317.Enabled = True
                Me.Comando363.Enabled = False
            Else
                Me.Comando330.Enabled = True
                Me.Comando366.Enabled = False
                Me.Comando317.Enabled = False
                Me.Comando363.Enabled = False
            End If
        End If
    Case Is < numerosemana(Date) - 1
        'Cuando se explora en dias de semanas 2 hacia atras
        Me.Comando330.Enabled = False
        Me.Comando366.Enabled = False
        Me.Comando317.Enabled = False
        Me.Comando363.Enabled = False
End Select

End Sub
