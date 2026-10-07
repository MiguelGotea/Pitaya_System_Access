' ==========================================================
' Modulo  : Form_OrdenDeCompraPlantilla
' Tipo    : 100  |  Lineas: 212
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database
Public Sub actualizarmodopagoinfo()
Select Case Me.tipopago

Case "TRANSFERENCIA"
    Me.reembolsodirecto.Visible = False
    
Case "EFECTIVO"
    Me.reembolsodirecto.Visible = True

Case "TC BAC"
    Me.reembolsodirecto = "TARJETA DE CREDITO BAC"
    Me.reembolsodirecto.Visible = True

Case "TC FICOHSA"
    Me.reembolsodirecto = "TARJETA DE CREDITO FICOHSA"
    Me.reembolsodirecto.Visible = True
    
Case "TC PRICESMART"
    Me.reembolsodirecto = "TARJETA DE CREDITO PRICESMART"
    Me.reembolsodirecto.Visible = True
    
Case Else
    Me.reembolsodirecto.Visible = False
    
End Select

End Sub

Public Sub guardarordendecompra()
'guardar orden de compra en imagen
Set Application.Printer = Application.Printers(elegirimpresora("ImpresoraConvertirImagen"))
[Forms]![OrdenDeCompraPlantilla].Form.SetFocus
[Forms]![OrdenDeCompraPlantilla].Form.Requery
[Forms]![OrdenDeCompraPlantilla].Form.Printer.Orientation = acPRORPortrait
[Forms]![OrdenDeCompraPlantilla].Form.Printer.PaperSize = acPRPSLetter
[Forms]![OrdenDeCompraPlantilla].Form.Printer.LeftMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.RightMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.BottomMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.TopMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.SetFocus

DoCmd.PrintOut , 1, 1

Set Application.Printer = Application.Printers(elegirimpresora("ImpresoraTintaOficinas"))

'Carpeta para archivo
Dim provee As String

provee = DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.aproovedor)
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Ordenes de Compra\" & provee

If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

Sleep (5000)

'copiar archivo a otra carpeta
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\imagentemporal.png"
h = Direccion & "\" & Me.ordenbusqueda & ".png"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

DoCmd.Close acForm, "OrdenDeCompraPlantilla"
End Sub
Public Sub autorizaroc()
Me.firmagerente.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Firmas\Gerente de Operaciones.png"
Me.firmacompras.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Firmas\Jefe de Compras.png"

End Sub
Private Sub CodigoBusqueda_Exit(Cancel As Integer)
Me.Requery
End Sub


Private Sub afechaorden_Exit(Cancel As Integer)
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE OrdenDeCompra SET OrdenDeCompra.fechaorden = #" & Me.afechaorden & "#" & _
" WHERE OrdenDeCompra.codordendecompra = " & Me.ordenbusqueda
DoCmd.SetWarnings True
End Sub

Private Sub cantidadorden_Exit(Cancel As Integer)
Me.Requery
End Sub

Private Sub Comando106_Click()
DoCmd.OpenForm "IngresoAutomaticoProductos", acNormal
[Forms]![IngresoAutomaticoProductos]![adestino] = "[SubOrdenDeCompra]"
[Forms]![IngresoAutomaticoProductos]![apreingreso] = Me.ordenbusqueda
[Forms]![IngresoAutomaticoProductos]![adesdeform] = "OrdenDeCompraPlantilla"
End Sub

Private Sub Comando1078_Click()
If MsgBox("Desea mandar a pagar orden de compra a telegram?", vbYesNo, "Confirmacion") = vbYes Then
    Call autorizaroc
    Call guardarordendecompra
    
    Dim capti As String
    Dim direc As String
    Dim archivo As String
    
    direc = "C:\Users\" & NombreSistema() & "\Google Drive BP\Compras\Ordenes de Compra\" & Me.anombre
    capti = Me.anombre & " , " & Me.tipopago & " , " & Me.reembolsodirecto & " , " & Me.totaltotale
    
    'enviar archivo por telegram
    archivo = Me.ordenbusqueda & ".png"
    'MsgBox capti
    
    Me.Requery
    Call EnviarImagenTelegram(direc & "\", archivo, capti, grupotcontabilidad())

    MsgBox "Mensaje enviado"
End If
End Sub

Private Sub Comando683_Click()
'If MsgBox("Desea imprimir Copia de Almacen?", vbYesNo, "COPIA DE ALMACEN") = vbYes Then
'    [Forms]![OrdenDeCompraPlantilla]![ocultar4].Visible = False
'    [Forms]![OrdenDeCompraPlantilla]![ocultar4].BorderStyle = 0
'    [Forms]![OrdenDeCompraPlantilla]![ocultar5].Visible = False
'    [Forms]![OrdenDeCompraPlantilla]![ocultar5].BorderStyle = 0
'    [Forms]![OrdenDeCompraPlantilla]![ocultar2].Visible = False
'    [Forms]![OrdenDeCompraPlantilla]![ocultar2].BorderStyle = 0
'End If

[Forms]![OrdenDeCompraPlantilla].Form.SetFocus
[Forms]![OrdenDeCompraPlantilla].Form.Requery
[Forms]![OrdenDeCompraPlantilla].Form.Printer.Orientation = acPRORLandscape
[Forms]![OrdenDeCompraPlantilla].Form.Printer.PaperSize = acPRPSLetter
[Forms]![OrdenDeCompraPlantilla].Form.Printer.LeftMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.RightMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.BottomMargin = 0
[Forms]![OrdenDeCompraPlantilla].Form.Printer.TopMargin = 0
DoCmd.PrintOut , 1, 1

'Call guardarordendecompra  'se oculta porque interrumpe la impresion
End Sub

Private Sub Comando685_Click()
Dim provee As String
Dim direc As String
Dim numtel As String
Dim mensa As String
Dim archivo As String

provee = DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.aproovedor)
direc = "C:\Users\" & NombreSistema() & "\Google Drive BP\Compras\Ordenes de Compra\" & provee
numtel = DLookup("[Numero]", "[Proovedores]", "[CodProovedor]=" & Me.aproovedor)

Call guardarordendecompra

'enviar archivo por wsp

mensa = primersaludo() & ", Adjunto Orden de Compra "
archivo = direc & "\" & Me.ordenbusqueda & ".png"

Call EnviarImagenWhatsapp(mensa, archivo, numtel)

MsgBox "Mensaje enviado"
End Sub

Private Sub Comando820_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM SubOrdenDeCompra WHERE codsubordendecompra = " & Me.codsubordendecompra
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Comando946_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "OrdenDeCompraPlantilla"

End Sub

Private Sub costounitario_Exit(Cancel As Integer)
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.logo.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Logo\Logo Claro.png"


Me.Comando820.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Eliminar.png"
Me.Comando106.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Nuevo.png"
Call actualizarmodopagoinfo
Me.InsideHeight = 14000
End Sub



Private Sub tipopago_Exit(Cancel As Integer)
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE OrdenDeCompra SET OrdenDeCompra.tipopago = '" & Me.tipopago & "'" & _
" WHERE OrdenDeCompra.codordendecompra = " & Me.ordenbusqueda
DoCmd.SetWarnings True

Call actualizarmodopagoinfo

If Me.tipopago = "EFECTIVO" Then
    Me.reembolsodirecto = "REEMBOLSO MIGUEL"
    Me.reembolsodirecto.SetFocus
End If
End Sub
