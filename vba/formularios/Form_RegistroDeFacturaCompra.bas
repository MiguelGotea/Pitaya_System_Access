' ==========================================================
' Modulo  : Form_RegistroDeFacturaCompra
' Tipo    : 100
' Lineas  : 507
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Public Sub bloquearedicionfactura()
Me.afechaorden.Enabled = False
Me.aoperario.Enabled = False
Me.aproovedor.Enabled = False
Me.afactura.Enabled = False
Me.Comando924.Enabled = False
Me.atipopago.Enabled = False
Me.Comando839.Enabled = False
Me.Comando851.Enabled = False
Me.cantidadfactura.Enabled = False
Me.costototalfactura.Enabled = False
Me.Destino.Enabled = False
Me.Comando106.Enabled = False
Me.Comando1050.Visible = True
Me.Comando1056.Visible = False
Me.Comando1058.Visible = False
End Sub

Public Sub desbloquearedicionfactura()
'Me.afechaorden.Enabled = True
'Me.aoperario.Enabled = True
Me.aproovedor.Enabled = True
Me.afactura.Enabled = True
Me.Comando924.Enabled = True
'Me.atipopago.Enabled = True
'Me.Comando839.Enabled = True
Me.Comando851.Enabled = True
'Me.cantidadfactura.Enabled = True
Me.costototalfactura.Enabled = True
Me.Destino.Enabled = True
'Me.Comando106.Enabled = True
'Me.Comando1050.Visible = False
Me.Comando1056.Visible = True
Me.Comando1058.Visible = False

If Me.atipopago = "CAJA" Then
    Me.costototalfactura.Enabled = False
End If
If Me.afechaorden < Date - 45 Then
    Call bloquearedicionfactura
End If
End Sub

Public Sub superdesbloquearedicionfactura()
Me.afechaorden.Enabled = True
Me.aoperario.Enabled = True
Me.aproovedor.Enabled = True
Me.afactura.Enabled = True
Me.Comando924.Enabled = True
Me.atipopago.Enabled = True
Me.Comando839.Enabled = True
Me.Comando851.Enabled = True
Me.cantidadfactura.Enabled = True
Me.costototalfactura.Enabled = True
Me.Destino.Enabled = True
Me.Comando106.Enabled = True
Me.Comando1050.Visible = True
Me.Comando1056.Visible = True
Me.Comando1058.Visible = True

End Sub

Function verificardatosllenos() As Integer
If IsNull(Me.aoperario) Or Me.aoperario = "" Then
    MsgBox "Por favor, ingrese el operario.", vbExclamation
    verificardatosllenos = 0
    Exit Function
ElseIf IsNull(Me.afactura) Or Me.afactura = "" Then
    MsgBox "Por favor, ingrese el número de factura.", vbExclamation
    verificardatosllenos = 0
    Exit Function
ElseIf IsNull(Me.aproovedor) Or Me.aproovedor = "" Then
    MsgBox "Por favor, ingrese el proveedor.", vbExclamation
    verificardatosllenos = 0
    Exit Function
ElseIf IsNull(Me.atipopago) Or Me.atipopago = "" Then
    MsgBox "Por favor, ingrese el tipo de pago.", vbExclamation
    verificardatosllenos = 0
    Exit Function
End If
verificardatosllenos = 1

End Function

Private Sub CodigoBusqueda_Exit(Cancel As Integer)
Me.Requery
End Sub

Private Sub afactura_Exit(Cancel As Integer)
If Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If
End Sub

Private Sub afechaorden_Exit(Cancel As Integer)
If Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If
End Sub

Private Sub aoperario_Exit(Cancel As Integer)
If Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If
End Sub

Private Sub aproovedor_Exit(Cancel As Integer)
If Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If
End Sub

Private Sub atipopago_Exit(Cancel As Integer)
If Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If
End Sub

Private Sub cantidadfactura_Exit(Cancel As Integer)
'If IsNull(Me.cantidadfactura) Then
'    'No hace nada
'    Me.cantidadfactura.SetFocus
'Else

Me.Requery
'End If
End Sub

Private Sub Comando1050_Click()
DoCmd.OpenForm "LogueoAutorizacion"
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(8, 18, Date)
[Forms]![LogueoAutorizacion]![CodigoBusqueda].Enabled = False
[Forms]![LogueoAutorizacion]![origenlogueo] = "RegistroDeFacturaCompra"
End Sub

Private Sub Comando1056_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "RegistroDeFacturaCompra"
End Sub

Private Sub Comando1058_Click()
DoCmd.OpenForm "Historial Ingresos", , , "[Fecha]=#" & Me.afechaorden & "#"
[Forms]![Historial Ingresos]![Comando95].Visible = True

End Sub

Private Sub costototalfactura_Exit(Cancel As Integer)
'If Me.costototalfactura = 0 Then
'    'No hace nada
'    'MsgBox "No"
'    Me.costototalfactura.SetFocus
'Else
'
    Me.Requery
'End If
End Sub

Private Sub Comando106_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM compras WHERE CodIngresoAlmacen = " & Me.CodIngresoAlmacen
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Comando839_Click()
' Verificar si todos los campos están llenos
If verificardatosllenos = 0 Then
    Exit Sub
End If

If Me.condicion = 1 And Me.tipofacturaform = "EDICION" Then
    Me.condicion = 0
End If

DoCmd.OpenForm "IngresoAutomaticoProductos", acNormal
[Forms]![IngresoAutomaticoProductos]![adestino] = "[Compras]"
[Forms]![IngresoAutomaticoProductos]![adesdeform] = "[RegistroDeFacturaCompra]"
[Forms]![IngresoAutomaticoProductos]![compraoperario] = Me.aoperario
[Forms]![IngresoAutomaticoProductos]![comprafactura] = Me.afactura
[Forms]![IngresoAutomaticoProductos]![compraproovedor] = Me.aproovedor
[Forms]![IngresoAutomaticoProductos]![compratipopago] = Me.atipopago
[Forms]![IngresoAutomaticoProductos]![afechapro] = Me.afechaorden

If codigoLocal <> 0 Then
   
    [Forms]![IngresoAutomaticoProductos]![Etiqueta153].Visible = False
    [Forms]![IngresoAutomaticoProductos]![aproducto].Visible = False
    
    [Forms]![IngresoAutomaticoProductos].Form.Filter = "[CompraDirectaSucursal] <>0 And [SubProducto]=0"
    [Forms]![IngresoAutomaticoProductos].Form.FilterOn = True
Else
    [Forms]![IngresoAutomaticoProductos]![modulo] = "contabilidad"
End If
End Sub



Private Sub Comando924_Click()
Me.condicion.Value = 0
' Verificar si el campo del proveedor está lleno
    If IsNull(Me.aproovedor) Or Me.aproovedor = "" Then
        MsgBox "Por favor, ingrese el proveedor antes de generar la factura.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la fecha y hora actual
    Dim fechaHoraActual As Date
    fechaHoraActual = Now

    ' Formatear la fecha y la hora según el formato deseado
    Dim formatoFecha As String
    formatoFecha = Format(fechaHoraActual, "ddmmyyhhnnss")
    
    ' Obtener las siglas del proveedor (primeras 2 letras de cada palabra)
    Dim siglasProveedor As String
    siglasProveedor = SiglasInicialesTexto(DLookup("[Nombre]", "[Proovedores]", "[CodProovedor] = " & Me.aproovedor), 2)

    ' Generar el número de factura
    Dim NumeroFactura As String
    NumeroFactura = formatoFecha & siglasProveedor

    ' Asignar el número de factura al cuadro de texto afactura
    Me.afactura = NumeroFactura

Me.afactura.Requery

End Sub


Private Sub Comando934_Click()

' Verificar si el campo afactura es nulo o vacío
   If IsNull(Me.afactura.Value) Or Me.afactura.Value = "" Then
       MsgBox "No hay nada que editar. El campo afactura está vacío.", vbExclamation
   Else
       ' Obtener el número actual de factura
       Dim numeroActual As String
       numeroActual = Me.afactura.Value
       
       ' Solicitar al usuario el nuevo número de factura
       Dim nuevoNumero As String
       nuevoNumero = InputBox("Ingrese el nuevo número de factura:", "Editar Número de Factura", numeroActual)

       ' Verificar si el usuario ingresó un nuevo número
       If nuevoNumero <> "" Then
           DoCmd.SetWarnings False
           ' Realizar la actualización en la tabla "Compras"
           DoCmd.RunSQL "UPDATE Compras SET NumeroFactura = '" & nuevoNumero & "' WHERE NumeroFactura = '" & numeroActual & "'"
           DoCmd.SetWarnings True
           ' Actualizar el formulario principal y el cuadro de texto con el nuevo número de factura
           Me.afactura.Value = nuevoNumero
       End If
   End If

End Sub

Private Sub Comando997_Click()

If Me.tipofacturaform = "NUEVO" Then

    '************ CASO FACTURA EN MODO NUEVO ********************

    If Me.condicion = 0 Then
        If comprobarProductosFactura(Me.afactura) = 0 Then
        ' No existen registros asociados al número de factura
        DoCmd.Close
        Else
        
            ' Comprobar si el CostoTotal es cero o nulo
            If costoTotalCeroNulo(Me.afactura) = 1 Then
                ' El CostoTotal es cero o nulo, mostrar mensaje y salir del proceso
                'MsgBox "Debe ingresar el costo total de los productos.", vbExclamation, "Error de Validación"
                If MsgBox("Existen productos con una cantidad total de 0, ¿desea salir?" & vbCrLf & vbCrLf & "Si: Borrar nueva factura" & vbCrLf & "No: Seguir editando factura", vbYesNo) = vbYes Then
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "DELETE * FROM Compras WHERE Compras.NumeroFactura = '" & Me.afactura & "'"
                    DoCmd.SetWarnings True
                    DoCmd.Close
                Else
                    'no hace nada
                End If
            Else
                ' Continuar con el proceso de guardado
                If MsgBox("¿Desea guardar la factura?", vbQuestion + vbYesNo, "Confirmar Guardar Cambios") = vbYes Then
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "UPDATE Compras SET Compras.Fecha = #" & Me.afechaorden & "#, " & _
                                 "Compras.CodProveedor = " & Me.aproovedor & ", " & _
                                 "Compras.Tipo = '" & Me.atipopago & "', " & _
                                 "Compras.CodOperario = " & Me.aoperario & ", " & _
                                 "Compras.NumeroFactura = '" & Me.afactura & "' " & _
                                 "WHERE Compras.NumeroFactura = '" & Me.facturainvisible & "'"
                    DoCmd.SetWarnings True
            
                    [Forms]![Ingreso de Compras].Form.Requery
                    Me.condicion = 1
                    DoCmd.Close
                Else
                    'no hace nada
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "DELETE * FROM Compras WHERE Compras.NumeroFactura = '" & Me.afactura & "'"
                    DoCmd.SetWarnings True
                    DoCmd.Close
                End If
            End If
        End If
    Else
        'no se ha modificado el numero de factura ni datos principales
        DoCmd.Close
    End If

Else '************ CASO FACTURA EN MODO EDICION ********************

    ' Comprobar si existen registros asociados al número de factura
    If comprobarProductosFactura(Me.facturainvisible) = 0 Then
        ' No existen registros asociados al número de factura
        DoCmd.Close
    Else
        If verificardatosllenos = 0 Then
            'no hace nada
        Else
            ' Comprobar si el CostoTotal es cero o nulo
            If costoTotalCeroNulo(Me.facturainvisible) = 1 Then
                ' El CostoTotal es cero o nulo, mostrar mensaje y salir del proceso
            
                MsgBox "Debe ingresar el costo total de los productos.", vbExclamation, "Error de Validación"
                
            Else
                ' El CostoTotal es diferente de cero y no nulo, preguntar si desea guardar la factura
                If Me.condicion = 0 Then
                    ' Continuar con el proceso de guardado
                    If MsgBox("¿Desea guardar la factura?", vbQuestion + vbYesNo, "Confirmar Guardar Cambios") = vbYes Then
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL "UPDATE Compras SET Compras.Fecha = #" & Me.afechaorden & "#, " & _
                                 "Compras.CodProveedor = " & Me.aproovedor & ", " & _
                                 "Compras.Tipo = '" & Me.atipopago & "', " & _
                                 "Compras.CodOperario = " & Me.aoperario & ", " & _
                                 "Compras.NumeroFactura = '" & Me.afactura & "' " & _
                                 "WHERE Compras.NumeroFactura = '" & Me.facturainvisible & "'"
                        DoCmd.SetWarnings True

                        [Forms]![Ingreso de Compras].Form.Requery
                        Me.condicion = 1
                        DoCmd.Close
                    Else
                        'no hace nada
                        DoCmd.Close
                    End If
                Else
                    'no se ha modificado el numero de factura ni datos principales
                    DoCmd.Close
                End If
            End If
        End If
    End If
End If

[Forms]![Ingreso de Compras].Form.Requery
End Sub

Private Sub Comando851_Click()

If IsNull(Me.afactura) Then
    'no hace nada
    MsgBox "Ingresar número de factura"
Else

    If Me.tipofacturaform = "NUEVO" Then
        
        '************ CASO FACTURA EN MODO NUEVO ********************
        
        ' Comprobar si existen registros asociados al número de factura
        If comprobarProductosFactura(Me.afactura) = 0 Then
            ' No existen registros asociados al número de factura
            ' No hace nada
        Else
            If verificardatosllenos = 0 Then
                'no hace nada
            Else
                ' Comprobar si el CostoTotal es cero o nulo
                If costoTotalCeroNulo(Me.afactura) = 1 Then
                    ' El CostoTotal es cero o nulo, mostrar mensaje y salir del proceso
                    MsgBox "Debe ingresar el costo total de los productos.", vbExclamation, "Error de Validación"
                Else
                    ' El CostoTotal es diferente de cero y no nulo, preguntar si desea guardar la factura
                    If Me.condicion = 0 Then
                        ' Continuar con el proceso de guardado
                        If MsgBox("¿Desea guardar la factura?", vbQuestion + vbYesNo, "Confirmar Guardar Cambios") = vbYes Then
                            DoCmd.SetWarnings False
                            DoCmd.RunSQL "UPDATE Compras SET Compras.Fecha = #" & Me.afechaorden & "#, " & _
                                     "Compras.CodProveedor = " & Me.aproovedor & ", " & _
                                     "Compras.Tipo = '" & Me.atipopago & "', " & _
                                     "Compras.CodOperario = " & Me.aoperario & ", " & _
                                     "Compras.NumeroFactura = '" & Me.afactura & "' " & _
                                     "WHERE Compras.NumeroFactura = '" & Me.facturainvisible & "'"
                            DoCmd.SetWarnings True
                            
                            [Forms]![Ingreso de Compras].Form.Requery
                            Me.condicion = 1
                            DoCmd.Close
                        Else
                            'no hace nada
                        End If
                        
                    Else
                        'no se ha modificado el numero de factura ni datos principales
                    End If
                End If
            End If
        End If
        
    Else ' ************ CASO FACTURA EN MODO EDICION ********************
    
        ' Comprobar si existen registros asociados al número de factura
        If comprobarProductosFactura(Me.facturainvisible) = 0 Then
            ' No existen registros asociados al número de factura
            ' No hace nada
        Else
            If verificardatosllenos = 0 Then
                'no hace nada
            Else
                ' Comprobar si el CostoTotal es cero o nulo
                If costoTotalCeroNulo(Me.facturainvisible) = 1 Then
                    ' El CostoTotal es cero o nulo, mostrar mensaje y salir del proceso
                    MsgBox "Debe ingresar el costo total de los productos.", vbExclamation, "Error de Validación"
                Else
                    ' El CostoTotal es diferente de cero y no nulo, preguntar si desea guardar la factura
                    If Me.condicion = 0 Then
                        ' Continuar con el proceso de guardado
                        If MsgBox("¿Desea guardar la factura?", vbQuestion + vbYesNo, "Confirmar Guardar Cambios") = vbYes Then
                            DoCmd.SetWarnings False
                            DoCmd.RunSQL "UPDATE Compras SET Compras.Fecha = #" & Me.afechaorden & "#, " & _
                                     "Compras.CodProveedor = " & Me.aproovedor & ", " & _
                                     "Compras.Tipo = '" & Me.atipopago & "', " & _
                                     "Compras.CodOperario = " & Me.aoperario & ", " & _
                                     "Compras.NumeroFactura = '" & Me.afactura & "' " & _
                                     "WHERE Compras.NumeroFactura = '" & Me.facturainvisible & "'"
                            DoCmd.SetWarnings True
                            
                            [Forms]![Ingreso de Compras].Form.Requery
                            Me.condicion = 1
                            DoCmd.Close
                        Else
                            'no hace nada
                        End If
                        
                    Else
                        'no se ha modificado el numero de factura ni datos principales
                    End If
                    MsgBox "Factura guardada"
                End If
                    
            End If
        End If
        
    End If
End If

End Sub

Function costoTotalCeroNulo(Factura As String) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Compras.CostoTotal, 1 AS Contador FROM Compras WHERE Compras.NumeroFactura = '" & Factura & "' AND (Compras.CostoTotal Is Null OR Compras.CostoTotal = 0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
costoTotalCeroNulo = rst("Contador")

rst.Close
Exit Function

Nulo:
costoTotalCeroNulo = 0
    
End Function

Function comprobarProductosFactura(Factura As String) As Integer
' Verifica si existen registros asociados al número de factura proporcionado
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Compras.CodCotizacion, 1 AS Contador FROM Compras WHERE Compras.NumeroFactura = '" & Factura & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
comprobarProductosFactura = rst("Contador")

rst.Close
Exit Function

Nulo:
comprobarProductosFactura = 0
End Function


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Comando106.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Eliminar.png"
If codigoLocal <> 0 Then
    Me.Lote.width = 0
    Me.Peso.width = 0
    Me.Destino.width = 0
End If
End Sub

