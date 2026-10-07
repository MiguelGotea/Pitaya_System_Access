' ==========================================================
' Modulo  : Form_RegistroPreIngresosPitaya
' Tipo    : 100
' Lineas  : 711
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Public Sub ingresocambiosprelista()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As Long
Dim afech As Date
Dim canti As Double
Dim cotglob As Integer
Dim codsubpor As Long
miSQL = "SELECT CambiosPreIngresosPitaya.CodPreIngresoPitaya, CambiosPreIngresosPitaya.CodCotizacion," & _
" CambiosPreIngresosPitaya.Cantidad, PreIngresoPitaya.Fecha, CambiosPreIngresosPitaya.CodCambiosPreIngresoPitaya" & _
" FROM CambiosPreIngresosPitaya" & _
" INNER JOIN PreIngresoPitaya ON CambiosPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" WHERE (((CambiosPreIngresosPitaya.CodPreIngresoPitaya)=" & Me.acodigo & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If rst.EOF Then
    ' no hay registros
    Exit Sub
End If
    
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = rst("Fecha")

For I = 1 To cantli
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        cotglob = PorcionGlobalDePorcion(cot)
        canti = rst("Cantidad")
        
        If cotglob <> 0 Then ' es porcion
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cotglob & ", " & canti & ", #" & afech & "#)"
            
            'crear subporcionamiento primero
            DoCmd.RunSQL "INSERT INTO SubPorcionamiento(Procedencia, CodProcesamiento, Cantidad, Fecha) values" & _
            " (" & cotglob & ", 0, " & canti & ", #" & afech & "#)"
            Me.Requery
            codsubpor = UltimoCodSubporcionamiento()
    
            'lugo crear porcionamiento normal
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia, CodSubPorcionamiento) values" & _
            " (" & cot & ", 0, " & canti & ", #" & afech & "#, " & OperarioAzar() & ", " & cotglob & ", " & codsubpor & ")"
            DoCmd.SetWarnings True
        Else 'cotiacion general
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cot & ", " & canti & ", #" & afech & "#)"
            DoCmd.SetWarnings True
        End If

    End If
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [StatusCambiosPreingreso](CodCambiosPreIngresoPitaya, Status, CodOperario, Fecha)" & _
    " values (" & rst("CodCambiosPreIngresoPitaya") & ", -1, " & Me.codigoope & " ,#" & Now & "#)"
    DoCmd.SetWarnings True
    
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call ingresocambiosprelista
End Sub

Public Sub ingresoprelista()

If MsgBox("Desea darle ingresos a todos los productos de la lista #" & Me.apreing & "?", vbYesNo, "Confirmacion") = vbNo Then
    Exit Sub
End If

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As Long
Dim afech As Date
Dim canti As Double
Dim cotglob As Integer
Dim codsubpor As Long

miSQL = "SELECT SubPreIngresosPitaya.CodPreIngresoPitaya, SubPreIngresosPitaya.CodCotizacion," & _
" SubPreIngresosPitaya.Cantidad, PreIngresoPitaya.Fecha" & _
" FROM SubPreIngresosPitaya" & _
" INNER JOIN PreIngresoPitaya ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" WHERE (((SubPreIngresosPitaya.CodPreIngresoPitaya)=" & Me.apreing & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = rst("Fecha")

For I = 1 To cantli
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        cotglob = PorcionGlobalDePorcion(cot)
        canti = rst("Cantidad")
        
        If cotglob <> 0 Then ' es porcion
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cotglob & ", " & canti & ", #" & afech & "#)"
            
            'crear subporcionamiento primero
            DoCmd.RunSQL "INSERT INTO SubPorcionamiento(Procedencia, CodProcesamiento, Cantidad, Fecha) values" & _
            " (" & cotglob & ", 0, " & canti & ", #" & afech & "#)"
            Me.Requery
            codsubpor = UltimoCodSubporcionamiento()
    
            'lugo crear porcionamiento normal
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia, CodSubPorcionamiento) values" & _
            " (" & cot & ", 0, " & canti & ", #" & afech & "#, " & OperarioAzar() & ", " & cotglob & ", " & codsubpor & ")"
            DoCmd.SetWarnings True
        Else 'cotiacion general
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cot & ", " & canti & ", #" & afech & "#)"
            DoCmd.SetWarnings True
        End If

    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [StatusPreingreso](CodPreIngresoPitaya, Status, CodOperario, Fecha)" & _
" values (" & Me.acodigo & ", -1, " & Me.codigoope & ", #" & Now & "#)"
DoCmd.SetWarnings True

MsgBox "Ingresos Completos"
Me.Status = "INGRESADO"
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call ingresoprelista
End Sub


Private Sub Comando148_Click()  ' borrar linea y sumezclas
On Error GoTo ErrorHandler
Dim cantip As Integer
llave = Me.CodSubPreIngresoPitaya

' Confirmar antes de eliminar
If MsgBox("¿Estás seguro de querer eliminar el registro?", vbYesNo + vbInformation, "Confirmar") = vbYes Then

    
    ' Comprobar si el producto es una mezcla

    If ExisteMezcla(Me.CodCotizacion) = 1 Then
        ' Si es una mezcla, eliminar la mezcla y sus subproductos
    
        ' Obtener la cantidad de subproductos
        cantip = cantidadProductosMezcla(Me.CodCotizacion)
    
        ' Eliminar subproductos relacionados de SubPreIngresosPitaya
        For sp = 1 To cantip
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM SubPreIngresosPitaya WHERE CodSubPreIngresoPitaya = " & llave + sp
            DoCmd.SetWarnings True
        Next sp
    End If
    
    ' Eliminar la mezcla original de SubPreIngresosPitaya
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM SubPreIngresosPitaya WHERE CodSubPreIngresoPitaya = " & llave
    DoCmd.SetWarnings True



Else
    'no hace nada
End If



' Actualizar el formulario o la lista para reflejar los cambios
Me.Requery

Exit Sub

ErrorHandler:
MsgBox "Error al eliminar el registro"
End Sub

Private Sub Comando270_Click() 'cambiar cantidad
On Error GoTo Nulo
Dim cantic As Double
Dim cantip As Integer
Dim llave As Long
Dim cantactual As Double
Dim cotix As Integer
Dim preinx As Long
Dim cotimezcla As Long
Dim nuevacanti As Double

preinx = Me.acodigo
llave = Me.CodSubPreIngresoPitaya
cantactual = Me.Cantidad
cotix = Me.CodCotizacion

If Me.tipovista = "modosucursal" Then ' cambio se refiere a correccion e ingresi como cambios
    If Me.Status = "INGRESADO" Then
        MsgBox "La lista ya ha sido ingreso, no se pueden hacer modificaciones"
        Exit Sub
    End If
    
    If Me.Cantidad < 0 Then
        MsgBox "No se permiten hacer cambios a los items de DEVOLUCION"
        Exit Sub
    End If
    
    cantic = InputBox("Ingresar cantidad real que ingreso", "ingreso Real", 0)
    If cantic < 0 Then
        MsgBox "no se permite ingresar una cantidad negativa"
        Exit Sub
    End If
    
    nuevacanti = cantic - (cantactual + cantidadcambiosproductopreingreso(cotix, llave))
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO CambiosPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya, CodSubPreIngresoPitaya)" & _
    " values (" & cotix & ", " & nuevacanti & ", " & preinx & ", " & llave & ")"
    DoCmd.SetWarnings True
    
    ' Comprobar si el producto es una mezcla
    
    If ExisteMezcla(cotix) = 1 Then
        ' Si es una mezcla, modificar la cantidad en la mezcla y sus subproductos
        cantip = cantidadProductosMezcla(cotix)
        ' Modificar la cantidad en subproductos relacionados de SubPreIngresosPitaya
            
        For sp = 1 To cantip
            cotimezcla = DLookup("[CodCotizacion]", "[SubPreIngresosPitaya]", "[CodSubPreIngresoPitaya]=" & llave + sp)
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO CambiosPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya, CodSubPreIngresoPitaya)" & _
            " values (" & cotimezcla & ", " & nuevacanti & ", " & preinx & ", " & llave + sp & ")"
            DoCmd.SetWarnings True
    
        Next sp
        
    End If
Else
    cantic = InputBox("Ingresar nueva cantidad", "Corregir Cantidad", 0)
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE SubPreIngresosPitaya SET Cantidad = " & cantic & " WHERE CodCotizacion = " & cotix & " AND CodPreIngresoPitaya = " & preinx
    DoCmd.SetWarnings True
    
    ' Comprobar si el producto es una mezcla
    
    If ExisteMezcla(cotix) = 1 Then
        ' Si es una mezcla, modificar la cantidad en la mezcla y sus subproductos
        cantip = cantidadProductosMezcla(cotix)
        ' Modificar la cantidad en subproductos relacionados de SubPreIngresosPitaya
            
        For sp = 1 To cantip
        
            DoCmd.SetWarnings False
            DoCmd.RunSQL "UPDATE SubPreIngresosPitaya SET Cantidad = " & cantic & " WHERE CodSubPreIngresoPitaya = " & llave + sp
            DoCmd.SetWarnings True
    
        Next sp
        
    End If
End If

Me.Requery
Exit Sub

Nulo:
MsgBox "Volver a intentar"

End Sub


Private Sub Comando214_Click()  ' Imprimir
If Me.validado = "VALIDADO" Then

    If ContadorItemsPreingreso(Me.CodPreIngresoPitaya) > 30 Then
    MsgBox "La lista tiene mas de 30 items, saldra una hoja adicional", vbCritical
    End If
    'copia de error
'    DoCmd.OpenReport "DetallePreingreso2", acViewReport, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
'    DoCmd.PrintOut acSelection, 1, 1
'    DoCmd.Close acReport, "DetallePreingreso2"
    
    DoCmd.OpenForm "DetallePreingreso2", acNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    '[Forms]![DetallePreingreso2].Form.Requery
    DoCmd.PrintOut
    DoCmd.Close acForm, "DetallePreingreso2"

Else
    MsgBox "Pendiente de validacion para poder imprimir la lista"
End If
End Sub



Private Sub Comando279_Click()  ' agregar cambios
DoCmd.OpenForm "RegistroCambiosPreIngresosPitaya", , , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
               " AND PorcionDentroDeMezcla([CambiosPreIngresosPitaya]![CodCotizacion])=0"

[Forms]![RegistroCambiosPreIngresosPitaya]![aencabezado].Caption = "REGISTRO DE CAMBIOS DE PREINGRESO " & UCase(DLookup("[Destino]", "[PreIngresoPitaya]", "[CodPreingresoPitaya]=" & Me.CodPreIngresoPitaya))
[Forms]![RegistroCambiosPreIngresosPitaya]![fechaac] = Me.fechaac
[Forms]![RegistroCambiosPreIngresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
[Forms]![RegistroCambiosPreIngresosPitaya]![codigoope] = Me.codigoope
End Sub

Private Sub Comando282_Click()  'generar ingresos de lista
If Me.fechaac <> Date Then
    MsgBox "No se pueden ingresar listas fuera de la fecha de recepcion"
    If MsgBox("Ingresar con permisos de administrador?", vbYesNo, "Permisos de Administrador") = vbYes Then
        DoCmd.OpenForm "IngresoClavePrivado"
        [Forms]![IngresoClavePrivado].Direccion = "IngresarPreingresoAdministrador"
    End If
    Exit Sub
End If

If Me.Status = "INGRESADO" Then
    MsgBox "La Lista ya ha sido ingresada al sistema"
    Exit Sub
End If

If Me.validado <> "VALIDADO" Then
    MsgBox "La lista no esta validada en la central, no se puede dar ingreso"
    Exit Sub
End If

Me.apreing = Me.CodPreIngresoPitaya
Call ingresoprelista
Call ingresocambiosprelista
[Forms]![HistorialPreIngresosLocal].Form.Requery

End Sub


'----------------  BOTONES EDICION DE PREINGRESOS

Private Sub Comando405_Click()  ' MOSTRAR Y DESPAARECES MEZCLAS

If Me.Comando405.Caption = "QUITAR FILTRO" Then
    Me.Filter = "[SubPreIngresosPitaya]![CodPreIngresoPitaya] = " & Me.acodigo
    Me.FilterOn = True
    Me.Comando405.Caption = "FILTRAR"

Else
    Me.Filter = "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.acodigo & _
               " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    Me.FilterOn = True
    Me.Comando405.Caption = "QUITAR FILTRO"
End If

End Sub


Private Sub Comando480_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (10, 0, " & Me.acodigo & ")"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (1146, 0, " & Me.acodigo & ")"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (1489, 0, " & Me.acodigo & ")"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (176, 0, " & Me.acodigo & ")"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (28, 0, " & Me.acodigo & ")"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (17, 0, " & Me.acodigo & ")"
DoCmd.SetWarnings True
Me.Requery
End Sub

Private Sub Comando500_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (746, 1, " & Me.acodigo & ")"
DoCmd.SetWarnings True
Me.Requery
End Sub


Private Sub Comando241_Click()
DoCmd.OpenForm "RegistrarProductosMostrador"
[Forms]![RegistrarProductosMostrador]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosMostrador]![adesdetabla] = "[SubPreIngresosPitaya]"
[Forms]![RegistrarProductosMostrador]![adesdeform] = "[RegistroPreIngresosPitaya]"
[Forms]![RegistrarProductosMostrador]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosMostrador]![guardado].width = 0
End Sub

Private Sub Comando257_Click()
DoCmd.OpenForm "RegistrarProductosNoPorciones"
[Forms]![RegistrarProductosNoPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosNoPorciones]![adesdetabla] = "[SubPreIngresosPitaya]"
[Forms]![RegistrarProductosNoPorciones]![adesdeform] = "[RegistroPreIngresosPitaya]"
[Forms]![RegistrarProductosNoPorciones]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosNoPorciones]![guardado].width = 0
End Sub

Private Sub Comando312_Click()
DoCmd.OpenForm "RegistrarProductosPorciones"
[Forms]![RegistrarProductosPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosPorciones]![adesdetabla] = "[SubPreIngresosPitaya]"
[Forms]![RegistrarProductosPorciones]![adesdeform] = "[RegistroPreIngresosPitaya]"
[Forms]![RegistrarProductosPorciones]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosPorciones]![guardado].width = 0
End Sub
Private Sub Comando517_Click()
DoCmd.OpenForm "RegistrarInsumosFijosConsumibles"
[Forms]![RegistrarInsumosFijosConsumibles]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarInsumosFijosConsumibles]![adesdetabla] = "[SubPreIngresosPitaya]"
[Forms]![RegistrarInsumosFijosConsumibles]![adesdeform] = "[RegistroPreIngresosPitaya]"
[Forms]![RegistrarInsumosFijosConsumibles]![apreingre] = Me.acodigo
[Forms]![RegistrarInsumosFijosConsumibles]![guardado].width = 0
End Sub

Private Sub Comando139_Click()
DoCmd.OpenForm "IngresoAutomaticoProductos", acNormal
[Forms]![IngresoAutomaticoProductos]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductos]![adestino] = "[SubPreIngresosPitaya]"
[Forms]![IngresoAutomaticoProductos]![apreingreso] = Me.acodigo
[Forms]![IngresoAutomaticoProductos]![adesdeform] = "[RegistroPreIngresosPitaya]"
End Sub


Public Sub modovistasucursal()
Me.tipovista = "modosucursal"
Me.[Comando214].Visible = False 'Imprimir

Me.[Comando312].Visible = False ' Ingresar prociones
Me.[Comando257].Visible = False ' Ingresar no porciones
Me.[Comando241].Visible = False ' Ingresar mostrador
Me.[Comando517].Visible = False ' ingrresar consumibles
Me.[Comando139].Visible = False ' ingresar otros productos
Me.[Comando405].Visible = False ' filtra no filtrar mexzclas
Me.TabCtl503.Visible = False 'Rotulo de yogurt y frutas locales automatico
Me.[Comando500].Visible = False ' ingresar 1 pote yogurt
Me.[Comando480].Visible = False ' ingresar frutas lcoales

Me.[Etiqueta593].Visible = True 'rotulo ingreso real
Me.[ingresoreal].Visible = True 'dato ingreso real
'Me.[ingresoreal].Width = 0 'ancho ingresoreal
Me.[Texto568].Locked = False 'dato de paquetes
Me.[Texto568].Enabled = False 'dato de paqeutes
Me.[Cantidad].Locked = False 'dato de cantida preingreso
Me.[Cantidad].Enabled = False 'dato de cantida preingreso

Me.[ajustes].width = 0 ' ancho de columna de ajustes

Me.[Texto568].width = 0 ' ancho de columna de paquetes
Me.[Cantidad].width = 0 ' ancho de columna de cantidad

Me.[Comando270].Visible = True 'boton cambiar cantidad
Me.[Comando148].Visible = False 'boton borar linea

Me.[Comando279].Visible = False 'Agregar cambios
Me.[Comando282].Visible = True 'Generar INgresos de Lista

Me.Etiqueta434.Visible = True 'etiqieta de stsatus
Me.Status.Visible = True ' dato de status

Me.Etiqueta618.Visible = False
Me.ingresorealsucursal.Visible = False
Me.ingresorealsucursal.width = 0.1

Me.Etiqueta557.Visible = False 'Validacion de impresion despacho
Me.validado.Visible = False ' valor de validacion impresion despacho
Me.Comando547.Visible = False ' boton cambiar validado o no
End Sub

Public Sub modovistasucursaladministracion()
Me.tipovista = "modosucursaladministracion"
Me.[Comando214].Visible = False 'Imprimir

Me.[Comando312].Visible = False ' Ingresar prociones
Me.[Comando257].Visible = False ' Ingresar no porciones
Me.[Comando241].Visible = False ' Ingresar mostrador
Me.[Comando517].Visible = False ' ingrresar consumibles
Me.[Comando139].Visible = False ' ingresar otros productos
Me.[Comando405].Visible = False ' filtra no filtrar mexzclas
Me.TabCtl503.Visible = False 'Rotulo de yogurt y frutas locales automatico
Me.[Comando500].Visible = False ' ingresar 1 pote yogurt
Me.[Comando480].Visible = False ' ingresar frutas lcoales

Me.[Etiqueta593].Visible = True 'rotulo ingreso real
Me.[ingresoreal].Visible = True 'dato ingreso real
'Me.[ingresoreal].Width = 0 'ancho ingresoreal
Me.[Texto568].Locked = False 'dato de paquetes
Me.[Texto568].Enabled = False 'dato de paqeutes
Me.[Cantidad].Locked = False 'dato de cantida preingreso
Me.[Cantidad].Enabled = False 'dato de cantida preingreso

'Me.[ajustes].Width = 0 ' ancho de columna de ajustes

Me.[Texto568].width = 0 ' ancho de columna de paquetes
Me.[Cantidad].width = 0 ' ancho de columna de cantidad

Me.[Comando270].Visible = True 'boton cambiar cantidad
Me.[Comando148].Visible = False 'boton borar linea

Me.[Comando279].Visible = False 'Agregar cambios
Me.[Comando282].Visible = True 'Generar INgresos de Lista

Me.Etiqueta434.Visible = True 'etiqieta de stsatus
Me.Status.Visible = True ' dato de status

Me.Etiqueta618.Visible = False
Me.ingresorealsucursal.Visible = False
Me.ingresorealsucursal.width = 0.1

Me.Etiqueta557.Visible = False 'Validacion de impresion despacho
Me.validado.Visible = False ' valor de validacion impresion despacho
Me.Comando547.Visible = False ' boton cambiar validado o no
End Sub

Public Sub modovistacentral()
Me.tipovista = "modocentral"
Me.[Comando214].Visible = False 'Imprimir

Me.[Comando312].Visible = False ' Ingresar prociones
Me.[Comando257].Visible = False ' Ingresar no porciones
Me.[Comando241].Visible = False ' Ingresar mostrador
Me.[Comando517].Visible = False ' ingrresar consumibles
Me.[Comando139].Visible = False ' ingresar otros productos
Me.[Comando405].Visible = False ' filtra no filtrar mexzclas
Me.TabCtl503.Visible = False 'Rotulo de yogurt y frutas locales automatico
Me.[Comando500].Visible = False ' ingresar 1 pote yogurt
Me.[Comando480].Visible = False ' ingresar frutas lcoales

Me.[Etiqueta593].Visible = False 'rotulo ingreso real
Me.[ingresoreal].Visible = False 'dato ingreso real
Me.[ingresoreal].width = 0 'ancho ingresoreal
Me.[Texto568].Locked = True 'dato de paquetes
Me.[Texto568].Enabled = True 'dato de paqeutes
Me.[Cantidad].Locked = True 'dato de cantida preingreso
Me.[Cantidad].Enabled = True 'dato de cantida preingreso

Me.[ajustes].width = 0 ' ancho de columna de ajustes

Me.[Comando270].Visible = False 'boton cambiar cantidad
Me.[Comando148].Visible = False 'boton borar linea

Me.[Comando279].Visible = False 'Agregar cambios
Me.[Comando282].Visible = False 'Generar INgresos de Lista

Me.Etiqueta434.Visible = False 'etiqieta de stsatus
Me.Status.Visible = False ' dato de status

Me.Etiqueta618.Visible = True
Me.ingresorealsucursal.Visible = True
'Me.ingresorealsucursal.Width = 0.1

Me.Etiqueta557.Visible = True 'Validacion de impresion despacho
Me.validado.Visible = True ' valor de validacion impresion despacho
Me.Comando547.Visible = False ' boton cambiar validado o no
End Sub

Public Sub modovistadespacho()
Me.tipovista = "mododespacho"
Me.[Comando214].Visible = False 'Imprimir

Me.[Comando312].Visible = True ' Ingresar prociones
Me.[Comando257].Visible = True ' Ingresar no porciones
Me.[Comando241].Visible = True ' Ingresar mostrador
Me.[Comando517].Visible = True ' ingrresar consumibles
Me.[Comando139].Visible = True ' ingresar otros productos
Me.[Comando405].Visible = True ' filtra no filtrar mexzclas
Me.TabCtl503.Visible = True 'Rotulo de yogurt y frutas locales automatico
Me.[Comando500].Visible = True ' ingresar 1 pote yogurt
Me.[Comando480].Visible = True ' ingresar frutas lcoales

Me.[Etiqueta593].Visible = False 'rotulo ingreso real
Me.[ingresoreal].Visible = False 'dato ingreso real
Me.[ingresoreal].width = 0 'ancho ingresoreal
Me.[Texto568].Locked = True 'dato de paquetes
Me.[Texto568].Enabled = True 'dato de paqeutes
Me.[Cantidad].Locked = True 'dato de cantida preingreso
Me.[Cantidad].Enabled = True 'dato de cantida preingreso

Me.[ajustes].width = 0 ' ancho de columna de ajustes

Me.[Comando270].Visible = True 'boton cambiar cantidad
Me.[Comando148].Visible = True 'boton borar linea

Me.[Comando279].Visible = False 'Agregar cambios
Me.[Comando282].Visible = False 'Generar INgresos de Lista

Me.Etiqueta434.Visible = False 'etiqieta de stsatus
Me.Status.Visible = False ' dato de status

Me.Etiqueta618.Visible = False
Me.ingresorealsucursal.Visible = False
Me.ingresorealsucursal.width = 0.1

Me.Etiqueta557.Visible = True 'Validacion de impresion despacho
Me.validado.Visible = True ' valor de validacion impresion despacho
Me.Comando547.Visible = True ' boton cambiar validado o no
End Sub

Private Sub Comando547_Click() 'validar preingreso


If MsgBox("Desea validar el preingreso y aprobar su impresion?", vbYesNo, "Validacion de Preingreso") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE PreIngresoPitaya SET PreIngresoPitaya.Validado = -1" & _
    " WHERE PreIngresoPitaya.CodPreIngresoPitaya = " & Me.acodigo
    DoCmd.SetWarnings True
    Me.validado = "VALIDADO"
    Me.validado.SetFocus
    Call modovistacentral
    Me.Comando214.Visible = True
    Call SyncKardexCentralDespachoCierre30Dias
End If
End Sub

Private Sub Comando615_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
" values (1327, 0, " & Me.acodigo & ")"
DoCmd.SetWarnings True
End Sub
Private Sub validarcambios()
'On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cante As Integer

Dim cantix As Double
Dim cantixanterior As Double
Dim cotix As Long
Dim subpreingrex As Long
Dim preingrex As Long

preingrex = Me.acodigo

miSQL = "SELECT CambiosPreingresoSucursalMixed.CodPreIngresoPitaya, CambiosPreingresoSucursalMixed.CodCotizacion," & _
" CambiosPreingresoSucursalMixed.Cantidad, CambiosPreingresoSucursalMixed.CodSubPreIngresoPitaya" & _
" FROM CambiosPreingresoSucursalMixed" & _
" WHERE (((CambiosPreingresoSucursalMixed.CodPreIngresoPitaya)=" & preingrex & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cante = rst.RecordCount
rst.MoveFirst

For I = 1 To cante
    If I >= Me.icoti Then
        cantix = rst("Cantidad")
        cotix = rst("CodCotizacion")
        subpreingrex = rst("CodSubPreIngresoPitaya")
        cantixanterior = DLookup("[Cantidad]", "[SubPreIngresosPitaya]", "[CodSubPreIngresoPitaya]=" & subpreingrex)
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE SubPreIngresosPitaya SET Cantidad = " & cantixanterior + cantix & "" & _
        " WHERE CodSubPreIngresoPitaya = " & subpreingrex
        DoCmd.SetWarnings True
    End If
    rst.MoveNext
Next I

rst.Close

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO AprobacionCambiosPreIngresoSucursal(CodPreIngresoPitaya, Status, Registro)" & _
" values (" & preingrex & ", -1, #" & Now() & "#)"
DoCmd.SetWarnings True


Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call validarcambios
End Sub
Private Sub Comando633_Click()
If MsgBox("Validar cambios registrados en la sucursal", vbYesNo, "Confirmacion") = vbYes Then
    Call validarcambios
    MsgBox "Cambios Validados y preingreso corregido"
    Me.statuscambios = -1
    Me.Comando633.Enabled = False
    
    Me.Requery
    [Forms]![CambiosGlobalRegistrados].Form.Requery
End If
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

