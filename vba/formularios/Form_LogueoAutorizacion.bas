' ==========================================================
' Modulo  : Form_LogueoAutorizacion
' Tipo    : 100
' Lineas  : 86
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database


Private Sub CodigoBusqueda_Exit(Cancel As Integer)
If Me.origenlogueo = "Pago Personal" And Me.CodigoBusqueda = cargooperariooperativo(5, codigoLocal(), Date) Then
    MsgBox "Usuario no permitido par esta funcion"
    Me.CodigoBusqueda = ""
End If
End Sub

Private Sub CodigoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Comando48_Click()
Dim userx As Integer


If IsNull(Me.CodigoBusqueda) Or Me.CodigoBusqueda = 0 Then
    'no hace nada
    MsgBox "Colaborador no puede estar vacio"
Else
    
    If IsNull(Me.password) Or Me.password = 0 Then
        'no hace nada
        MsgBox "Ingresar contraseña"
    Else
        userx = Me.CodigoBusqueda

        If Me.password = CorroborarClave(Me.CodigoBusqueda) Or Me.password = CorroborarClave(cargooperariooperativo(15, 18, Date)) Then
            Select Case Me.origenlogueo
                Case "[Introduccion]"
                    'Call EnviarMensajeTelegram("Abriendo Facturacion '" & NombreOperario(userx) & "'", grupotgerencia())
                    DoCmd.OpenForm "Main Pitaya"
                    [Forms]![Main Pitaya]![codigologin] = userx
                    [Forms]![Main Pitaya].Form.Requery
                    DoCmd.Close acForm, "Introduccion"

                Case "DesbloquearInventario"
                    Call Forms("[Ingreso Inventario Pitaya]").bloquearfecha
                    Call Forms("[Ingreso Inventario Pitaya]").cargarllaveatemporal
                    Call Forms("[Ingreso Inventario Pitaya]").ocultartodoventanainventario
                    'Call Forms("[Ingreso Inventario Pitaya]").estadoBotones
                    
                    [Forms]![Ingreso Inventario Pitaya].Estado = "B"
                    [Forms]![Ingreso Inventario Pitaya].Comando330.Enabled = False
                    [Forms]![Ingreso Inventario Pitaya].Comando366.Enabled = False
                    [Forms]![Ingreso Inventario Pitaya].Comando317.Enabled = False
                    [Forms]![Ingreso Inventario Pitaya].Comando363.Enabled = True
                    [Forms]![Ingreso Inventario Pitaya].codigologin = userx
                    
                Case "DesbloquearHorarioSucursal"
                    Call Forms("[RegistroHorariosOperariosSemana]").modohorarionobloqueado
                    
                Case "RegistroDeFacturaCompra"
                    Call Forms("RegistroDeFacturaCompra").desbloquearedicionfactura
                    
                Case "IngresoAjustesInventario"
                    If Me.CodigoBusqueda = 321 Or Me.CodigoBusqueda = 373 Or Me.CodigoBusqueda = 366 Or Me.CodigoBusqueda = 808 Then
                        DoCmd.OpenForm Me.origenlogueo
                        Forms(Me.origenlogueo).Controls("codigologin") = userx
                    End If
                  
                Case Else ' resto de formularios
                    DoCmd.OpenForm Me.origenlogueo
                    Forms(Me.origenlogueo).Controls("codigologin") = userx
                    
            End Select
            
            DoCmd.Close acForm, "LogueoAutorizacion"
            
        Else
            MsgBox "Clave Erronea", vbOKOnly, "CLAVE ERRONEA"
            Me.password = ""
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

