' ==========================================================
' Modulo  : Form_IngresarClienteDelivery
' Tipo    : 100  |  Lineas: 166
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database

Private Sub aCodLocal_Exit(Cancel As Integer)
Me.aCodProovedoresDelivery.Requery
End Sub
Private Sub aCodProovedoresDelivery_Exit(Cancel As Integer)
If DLookup("[Nombre]", "[ProovedoresDelivery]", "[CodProovedoresDelivery]=" & Me.aCodProovedoresDelivery) = "Motorizado Pitaya" Then
    Me.Etiqueta129.Visible = True
    Me.aDistancia.Visible = True
    Me.Etiqueta121.Visible = True
    Me.aCodigoMotorizado.Visible = True
Else
    Me.Etiqueta129.Visible = False
    Me.aDistancia.Visible = False
    Me.aDistancia = 0
    Me.Etiqueta121.Visible = False
    Me.aCodigoMotorizado.Visible = False
    Me.aCodigoMotorizado = ""
End If
End Sub

Private Sub aTipoPedido_Exit(Cancel As Integer)
If Me.aTipoPedido = "Delivery" Then
    Me.Etiqueta9.Visible = True
    Me.aCodProovedoresDelivery.Visible = True
    Me.Etiqueta129.Visible = True
    Me.aDistancia.Visible = True
    Me.Etiqueta121.Visible = True
    Me.aCodigoMotorizado.Visible = True
Else
    Me.Etiqueta9.Visible = False
    Me.aCodProovedoresDelivery.Visible = False
    Me.aCodProovedoresDelivery = ""
    Me.Etiqueta129.Visible = False
    Me.aDistancia.Visible = False
    Me.aDistancia = 0
    Me.Etiqueta121.Visible = False
    Me.aCodigoMotorizado.Visible = False
    Me.aCodigoMotorizado = ""
End If
End Sub

Private Sub Comando170_Click()
Me.kmcalculados = CalcularDistanciaKM(Me.udesde, Me.uhasta)

DoCmd.SetWarnings False
Application.FollowHyperlink GenerarURLGoogleMapsCompleto(Me.udesde, Me.uhasta, "optimal")
DoCmd.SetWarnings True
End Sub

Private Sub Comando36_Click()

If Me.ACodLocal = "" Or IsNull(Me.ACodLocal) Then
    MsgBox "Ingresar sucursal a la que destinara el pedido"
    Exit Sub
End If

If Me.aTelefono = "" Or IsNull(Me.aTelefono) Then
    MsgBox "Ingresar un telefono del cliente "
    Exit Sub
End If

If Me.aDireccion = "" Or IsNull(Me.aDireccion) Then
    MsgBox "Ingresar la direccion de envio"
    Exit Sub
End If

If Me.aTipoPedido = "" Or IsNull(Me.aTipoPedido) Then
    MsgBox "Ingresar el tipo de servicio delivery"
    Exit Sub
End If

If Me.aTipoPedido = "Delivery" Then

    If Me.aCodProovedoresDelivery = "" Or IsNull(Me.aCodProovedoresDelivery) Then
        MsgBox "Ingresar proovedor de servicio delivery que se encargara"
        Exit Sub
    End If
    
    If DLookup("[Nombre]", "[ProovedoresDelivery]", "[CodProovedoresDelivery]=" & Me.aCodProovedoresDelivery) = "Motorizado Pitaya" Then
    
        If Me.aCodigoMotorizado = "" Or IsNull(Me.aCodigoMotorizado) Then
            MsgBox "Ingresar motorizado"
            Exit Sub
        Else
            [Forms]![Nota de Pedido]![CodMotorizado] = Me.aCodigoMotorizado
        End If
        
        If Me.aDistancia = "" Or IsNull(Me.aDistancia) Then
            MsgBox "Ingresar distancia"
            Exit Sub
        End If
    End If
End If





If Me.aCostoDelivery = "" Or IsNull(Me.aCostoDelivery) Then
    MsgBox "Ingresar costo del servicio"
    Exit Sub
End If

If Me.tiporegistro = 0 Then
    DoCmd.SetWarnings False
    If Me.aTipoPedido = "Delivery" Then
        If DLookup("[Nombre]", "[ProovedoresDelivery]", "[CodProovedoresDelivery]=" & Me.aCodProovedoresDelivery) = "Motorizado Pitaya" Then
            DoCmd.RunSQL "INSERT INTO ClientesDelivery(CodLocal, Telefono, Direccion, TipoPedido, CodProovedoresDelivery, CostoDelivery, CodigoMotorizado, Distancia)" & _
            " values (" & Me.ACodLocal & ", " & Me.aTelefono & ", '" & Me.aDireccion & "', '" & Me.aTipoPedido & "', " & Me.aCodProovedoresDelivery & ", " & Me.aCostoDelivery & ", " & Me.aCodigoMotorizado & ", " & Me.aDistancia & ")"
        Else 'otro seviio delivery
            DoCmd.RunSQL "INSERT INTO ClientesDelivery(CodLocal, Telefono, Direccion, TipoPedido, CodProovedoresDelivery, CostoDelivery, Distancia)" & _
            " values (" & Me.ACodLocal & ", " & Me.aTelefono & ", '" & Me.aDireccion & "', '" & Me.aTipoPedido & "', " & Me.aCodProovedoresDelivery & ", " & Me.aCostoDelivery & ", 0)"
        End If
    Else
        DoCmd.RunSQL "INSERT INTO ClientesDelivery(CodLocal, Telefono, Direccion, TipoPedido, CostoDelivery)" & _
        " values (" & Me.ACodLocal & ", " & Me.aTelefono & ", '" & Me.aDireccion & "', '" & Me.aTipoPedido & "', " & Me.aCostoDelivery & ")"
    End If
    DoCmd.SetWarnings True
    
    Dim ultimodeliv As Long
    ultimodeliv = ultimoclientedeliveryregistrado()
    
    [Forms]![Nota de Pedido]![CodClientesDelivery] = ultimodeliv
Else
    DoCmd.SetWarnings False
    If Me.aTipoPedido = "Delivery" Then
        If DLookup("[Nombre]", "[ProovedoresDelivery]", "[CodProovedoresDelivery]=" & Me.aCodProovedoresDelivery) = "Motorizado Pitaya" Then
            DoCmd.RunSQL "UPDATE ClientesDelivery SET" & _
            " ClientesDelivery.CodLocal = " & Me.ACodLocal & ", ClientesDelivery.Telefono = " & Me.aTelefono & "," & _
            " ClientesDelivery.Direccion = '" & Me.aDireccion & "', ClientesDelivery.TipoPedido = '" & Me.aTipoPedido & "'," & _
            " ClientesDelivery.CodProovedoresDelivery = " & Me.aCodProovedoresDelivery & ", ClientesDelivery.CostoDelivery = " & Me.aCostoDelivery & "," & _
            " ClientesDelivery.CodigoMotorizado = " & Me.aCodigoMotorizado & ", ClientesDelivery.Distancia = " & Me.aDistancia & _
            " WHERE ClientesDelivery.CodClientesDelivery =" & Me.codigoencurso
        Else 'otro servicio delivery
            DoCmd.RunSQL "UPDATE ClientesDelivery SET" & _
            " ClientesDelivery.CodLocal = " & Me.ACodLocal & ", ClientesDelivery.Telefono = " & Me.aTelefono & "," & _
            " ClientesDelivery.Direccion = '" & Me.aDireccion & "', ClientesDelivery.TipoPedido = '" & Me.aTipoPedido & "'," & _
            " ClientesDelivery.CodProovedoresDelivery = " & Me.aCodProovedoresDelivery & ", ClientesDelivery.CostoDelivery = " & Me.aCostoDelivery & "," & _
            " ClientesDelivery.CodigoMotorizado = null, ClientesDelivery.Distancia = 0" & _
            " WHERE ClientesDelivery.CodClientesDelivery =" & Me.codigoencurso
        End If
    Else 'retiro en local
        DoCmd.RunSQL "UPDATE ClientesDelivery SET" & _
        " ClientesDelivery.CodLocal = " & Me.ACodLocal & ", ClientesDelivery.Telefono = " & Me.aTelefono & "," & _
        " ClientesDelivery.Direccion = '" & Me.aDireccion & "', ClientesDelivery.TipoPedido = '" & Me.aTipoPedido & "'," & _
        " ClientesDelivery.CodProovedoresDelivery = null, ClientesDelivery.CostoDelivery = " & Me.aCostoDelivery & "," & _
        " ClientesDelivery.CodigoMotorizado = null, ClientesDelivery.Distancia = 0" & _
        " WHERE ClientesDelivery.CodClientesDelivery =" & Me.codigoencurso
    End If
    DoCmd.SetWarnings True
    
    [Forms]![Nota de Pedido]![CodClientesDelivery] = Me.codigoencurso
End If

'MsgBox "Datos de servicio delivery guardado correctamente"
DoCmd.Close acForm, "IngresarClienteDelivery"
End Sub

Private Sub Comando37_Click()
DoCmd.Close acForm, "IngresarClienteDelivery"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
