' ==========================================================
' Modulo  : Form_IngresoClavePrivado
' Tipo    : 100
' Lineas  : 121
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database

Private Sub botonbloquear_Click()

If Me.clavedesbloquear = CorroborarClave(5) Then
    Select Case Me.Direccion
        'Funciones especiales que no impliquen abrir un formualrio
        Case "desbloquearpedido"
            Call Forms("Nota de Pedido").DesbloquearFacturacion
            [Forms]![Nota de Pedido]![pacumulado] = [Forms]![Nota de Pedido]![piniciales] - [Forms]![Nota de Pedido]![ppedido]
        Case "IngresarPreingresoAdministrador"
            [Forms]![RegistroPreingresosPitaya]![apreing] = [Forms]![RegistroPreingresosPitaya]![CodPreIngresoPitaya]
            Call Forms("[RegistroPreingresosPitaya]").ingresoprelista
            [Forms]![RegistroPreingresosPitaya].Form.Requery
            Call Forms("[RegistroPreingresosPitaya]").ingresocambiosprelista
            [Forms]![HistorialPreIngresosLocal].Form.Requery
        Case "desbloquearhorario"
            Call Forms("[RegistroHorariosOperariosSemana]").modohorarionobloqueado
        Case "OrdenDeCompraPlantilla"
            Call Forms("[OrdenDeCompraPlantilla]").autorizaroc
            Call Forms("[OrdenDeCompraPlantilla]").guardarordendecompra
            'DoCmd.Close acForm, "OrdenDeCompraPlantilla"
        Case "actualizartablasmain"
            Call eliminartablasmain
            Call importartablasmain
            
            Call eliminartablascentral(codigoLocal())
            Call importartablascentral(codigoLocal())
            
            If APIDisponible() Then
                Call importartablasweb
            Else
                Call importartablaespecifica("RRHH", "Operarios", "Operarios", 3)
                Call importartablaespecifica("RRHH", "AsignacionNivelesCargos", "AsignacionNivelesCargos", 3)
                Call importartablaespecifica("RRHH", "NivelesCargos", "NivelesCargos", 3)
            End If
            MsgBox "Tablas Main Actualizadas"
        Case "actualizartablasmixedglobal"
            Call actualizartablasmixedglobal
            MsgBox "Tablas de Sucursales actualizadas"
        Case "GestionSistema"
            [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = True
            [Forms]![Menu Gestion]![main].Pages("Host").Visible = True
        Case "actualizardatosclub"
            Call actualizardatosclubmixedglobal
            'Call DescargarTablaCompleta("clientesclub", "clientesclubexterno", "sucursal <> " & codigolocal())
            MsgBox "Tablas de clientes actualizadas"
        Case "CierreCorroborarDatosPOS"
            Call Forms("CierreCorroborarDatosPOS").forzaringresopos
        Case "CierreCorroborarDatosDELIVERY"
            Call Forms("CierreCorroborarDatosDELIVERY").forzaringresopos
        Case "DesbloquearInventarioMain"
            Call Forms("Ingreso Inventario Pitaya").Comando317_Click
        Case "CierreCorroborarDatosTRANSFERENCIAS"
            Call Forms("[CierreCorroborarDatosTRANSFERENCIAS]").forzaringresopos
        Case "RegistroDeFacturaCompra"
            Call Forms("[RegistroDeFacturaCompra]").superdesbloquearedicionfactura
        Case "actualizartablaslocales"
            Dim busca As String
            Call eliminararchivosllavelectura
            busca = CurrentProject.Name
            
            If busca Like "Pitaya*" Then 'empieza con pitaya entonces si puede er sucursal o raiz\
                busca = Left(Right(busca, Len(busca) - 6), 1)
                If IsNumeric(busca) Then
                    indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
                    Call actualizartablas(3, CInt(indi))
                    MsgBox "Tablas locales vinculadas correctamente"
                    MsgBox ListarTablasVinculadasAgrupadasCompacto()
                End If
            Else
                MsgBox "Tiene que ser un archivo de pitaya local para cargar datos de local en edicion"
            End If
        Case "actualizartablasmodulo"
        
            Dim archim As String
            Dim esmodulo As String
            archim = CurrentProject.Name
            If archim Like "Modulo*" Then 'empieza con modulo entonces si puede er modulo \
                esmodulo = Split(Split(archim, "Modulo")(1), ".accdb")(0)
                If esmodulo = Me.variable Then 'modulo enviado es mismo nombre de archivo de modulo
                    Call actualizartablasmodulo(esmodulo)
                Else
                    MsgBox "No existe modulo seun nombre de archivo"
                End If
            Else
                MsgBox "Nombre de archivo no corresponde a ningun modulo"
            End If
            
        Case "desbloquearcompralocal"
            DoCmd.SetWarnings False
            DoCmd.RunSQL "UPDATE Cotizaciones SET compradirectasucursal = -1 WHERE CodCotizacion = " & Me.variable
            DoCmd.SetWarnings True
            
        Case Else
            ' cuando no se hace alguna funcion por defecto solo abre la ventana enviada como dato
            DoCmd.OpenForm Me.Direccion
    End Select
    
    DoCmd.Close acForm, "IngresoClavePrivado"
Else
    MsgBox "Clave Erronea"
    Me.clavedesbloquear = ""
    Me.clavedesbloquear.SetFocus
End If





End Sub

Private Sub Comando1253_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

