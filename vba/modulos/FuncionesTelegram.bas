' ==========================================================
' Modulo  : FuncionesTelegram
' Tipo    : 1
' Lineas  : 556
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database

Sub Test2()
'mas lento
Dim oXHTTP As Object
Dim url As String
Dim request As String
url = "https://api.telegram.org/bot5727507523:AAFqKB3YMkuQ9ocMb5nn2NWLwPvx-c34w8Q/getUpdates"

'borrarcache
url = url & "?cb=" & Timer() * 100
'MsgBox url
Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
oXHTTP.Open "GET", url, False
oXHTTP.send
'MsgBox oXHTTP.responseText
request = oXHTTP.responseText
Set oXHTTP = Nothing

End Sub
Sub ComunicacionTelegram()
On Error GoTo Nulo
'Habilitar Referencia Microsoft XML v6.0
'Descatvar privacidad de grupo de bot en telegram
Dim url As String
Const RunAsync As Boolean = True
Const ProcessComplete As Integer = 4

Dim request As MSXML2.XMLHTTP60
Set request = New MSXML2.XMLHTTP60

Dim response As String
Dim token As String

token = DLookup("[IdBotTe]", "DatosSistema")

url = "https://api.telegram.org/bot" & token & "/getUpdates"
url = url & "?cb=" & Timer() * 100
With request
    .Open "GET", url, RunAsync
    .setRequestHeader "Content-Type", "application/json"
    .send
    Do While request.ReadyState <> ProcessComplete
        DoEvents
    Loop
    response = .responseText
End With
Set request = Nothing

Dim nombrebot As String
nombrebot = DLookup("[BotTelegram]", "[StatusSucursales]", "[CodLocal]=" & codigoLocal())

If InStr(response, nombrebot) > 0 Then ' lo nombra en  el texto
    Dim divisionactu() As String
    Dim ccuenta As Integer
    Dim formu As String
    Dim ultimaactu As Long
    formu = Split(Split(response, nombrebot & Chr(32))(1), Chr(34))(0) 'formula + texto posterior competo
    Call formulastelegram(formu)
    
    divisionactu = Split(response, "update_id" & Chr(34) & ":")
    ccuenta = UBound(divisionactu) - LBound(divisionactu) + 1
    ultimaactu = Split(Split(response, "update_id" & Chr(34) & ":")(ccuenta - 1), ",")(0)
    Call LimpiarHistorialTelegram(ultimaactu)
Else
    'no existe ninguna indicacion para este bot, no hacer nada

End If

Exit Sub

Nulo:
'No hacer nada
End Sub

Sub formulastelegram(textocompleto As String)
On Error GoTo Nulo
Dim Detalle() As String
Detalle = Split(textocompleto, Chr(32))

Select Case Detalle(0)
        
    Case "foto"
        Call FotoTemporalTelegram
        
    Case "mensaje"
        Dim mensajesep As String
        Dim counte As Integer
        mensajesep = ""
        
        For counte = 1 To UBound(Detalle) - LBound(Detalle)
            mensajesep = mensajesep & " " & UCase(Detalle(counte))
        Next counte
        
        DoCmd.Close acForm, "Alarmas Programadas"
        DoCmd.OpenForm "Alarmas Programadas"
        Forms![ALarmas Programadas].Mensaje = mensajesep
        Call EnviarMensajeTelegram("Mensaje enviado", grupotoperaciones())
    
    Case "ventas"
    
        Call EnviarMensajeTelegram(CStr(AcumuladoDiaGuardado(Date)), grupotoperaciones())
        
    Case "desbloquear"
    
        DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Detalle(1)
        Call Forms("Nota de Pedido").DesbloquearFacturacion
        Call EnviarMensajeTelegram("Pedido " & CLng(Detalle(1)) & " Desbloqueado", grupotoperaciones())
        [Forms]![Nota de Pedido]![pacumulado] = [Forms]![Nota de Pedido]![piniciales] - [Forms]![Nota de Pedido]![ppedido]
    Case "consumo"

    Dim archivo As String
    Dim semanow As Integer
    Dim codin As String
    semanow = numerosemana(Date)
    
    codin = DLookup("[CodIngrediente]", "[DBIngredientes]", "[Nombre] like '" & Detalle(1) & "*'")
    
    If IsNull(codin) Then
        
        Call EnviarMensajeTelegram("No existe nombre de insumo", grupotoperaciones())
    
    Else
    
        archivo = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\consumoingredientetemporal.jpg"
        DoCmd.OpenForm "Dashboard_Consumos_Semanales_Insumo", acFormPivotChart, , "[semana]>=" & semanow - 8 & " AND [semana]<=" & semanow & " AND [CodIngrediente]='" & codin & "'"
        [Forms]![Dashboard_Consumos_Semanales_Insumo].Form.Requery
        [Forms]![Dashboard_Consumos_Semanales_Insumo].ChartSpace.exportpicture archivo
        DoCmd.Close acForm, "Dashboard_Consumos_Semanales_Insumo"
        
        Call EnviarImagenTelegram("C:\Users\" & NombreSistema() & "\Desktop\Sistema\", "consumoingredientetemporal.jpg", "Consumo Semanal", grupotoperaciones)
        
    End If

    Case "rustdesk"
    
        DoCmd.SetWarnings False
        Application.FollowHyperlink "C:\Program Files\RustDesk\rustdesk.exe"
        DoCmd.SetWarnings True
        
    Case "actualizardatos"
    
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
        
        'Call EnviarMensajeTelegram("Base de Datos Principal Actualizado", grupotoperaciones())
        MsgBox "Base de Datos Principal Actualizado"
        
    Case "actualizartablasmixed"
    
        Call actualizartablasmixedglobal
        
    Case "PedidoAprobado"
    
        'Call ReiniciarGoogleDrive
        'Call actualizartablasmixedglobal
    
    Case "reiniciardrive"
    
        Call ReiniciarGoogleDrive
        
        Call EnviarMensajeTelegram("Drive actualizado", grupotgerencia())
    Case "actualizardatosclub"
    
        Call actualizardatosclubmixedglobal
        'Call DescargarTablaCompleta("clientesclub", "clientesclubexterno", "sucursal <> " & codigolocal())
        
    Case "numerosemana"
    
        Call EnviarMensajeTelegram(str(numerosemana(Date)), grupotoperaciones())
        
    Case "descargardatossemanales"
        
        Call EnviarMensajeTelegram("Empezando a Descargar datos de la semana", grupotoperaciones())
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![semguar] = CLng(Detalle(1))
        Call Forms("[Menu Gestion]").Comando2352_Click
        Call EnviarMensajeTelegram("Datos semanales descargados completamente", grupotoperaciones())
        DoCmd.Close acForm, "Menu Gestion"
        
    Case "anular"
    
        Dim motivanu As String
        Dim counta As Integer
        motivanu = ""
        
        For counta = 2 To UBound(Detalle) - LBound(Detalle)
            motivanu = motivanu & " " & Detalle(counta)
        Next counta
        
        If motivosolicitudanulacionpedido(CLng(Detalle(1))) <> "" Then
            motivanu = motivosolicitudanulacionpedido(CLng(Detalle(1)))
        End If
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE NotaDePedido SET NotaDePedido.Anulado <>0, NotaDePedido.TotalGuardado = 0," & _
        " NotaDePedido.MotivoAnulado = '" & motivanu & "' WHERE NotaDePedido.CodPedido = " & CLng(Detalle(1))
        DoCmd.SetWarnings True
        
        If existesolicitudanulacionpedido(CLng(Detalle(1))) = 1 Then ' existe solicitud de anulacion
            DoCmd.SetWarnings False
            DoCmd.RunSQL "UPDATE AnulacionPedidos SET AnulacionPedidos.Status = 1, AnulacionPedidos.HoraAnulada = #" & Now & "#" & _
            " WHERE AnulacionPedidos.CodPedido = " & CLng(Detalle(1))
            DoCmd.SetWarnings True
        Else ' se solicita de la nada
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO AnulacionPedidos(CodPedido, HoraSolicitada, HoraAnulada, Status, Modalidad, CodPedidoCambio, Motivo)" & _
                         " values (" & CLng(Detalle(1)) & ", #" & Time & "#, #" & Time & "#,1,2, 0, '" & motivanu & "')"
            DoCmd.SetWarnings True
        End If
        
        Call EnviarMensajeTelegram("Pedido " & CLng(Detalle(1)) & " Anulado " & NombreSistema(), grupotanulaciones())
        
    Case "reiniciarcierre"
    
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE CierreDiario SET CierreDiario.HoraFinal = #" & DLookup("[HoraInicial]", "[CierreDiario]", "[CodigoCierre]=" & CLng(Detalle(1))) + 1 / 60 / 24 & "#" & _
        " WHERE CierreDiario.CodigoCierre = " & CLng(Detalle(1))
        DoCmd.SetWarnings True
        
        If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
            [Forms]![Main Pitaya].Form.Requery
        End If
        
        Call EnviarMensajeTelegram("Cierre " & CLng(Detalle(1)) & " anulado " & NombreSistema(), grupotanulaciones())
        
    Case "pedidosyacancelado"
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 106" & _
        " WHERE SubPedido.CodPedido = " & CLng(Detalle(1))
        DoCmd.SetWarnings True
        
        Call EnviarMensajeTelegram("Pedido " & CLng(Detalle(1)) & " reportado ya preparado sin devolucion de PedidosYa", grupotanulaciones())
        CreateObject("Shell.Application").Open DLookup("[FormularioCanceladoPedidosYa]", "[SistemaGlobal]")
    Case "turnos"
        
        Dim rst As DAO.Recordset
        Dim miSQL As String
        Dim ind As Integer
        Dim turnos As String
        turnos = ""
        ind = 1
        
        miSQL = "SELECT Operarios.Nombre, Operarios.Apellido, RegistroHorario.Ingreso, RegistroHorario.Salida, RegistroHorario.Fecha" & _
        " FROM RegistroHorario INNER JOIN Operarios ON RegistroHorario.CodOperario = Operarios.CodOperario" & _
        " WHERE (((RegistroHorario.Salida) Is Null) AND ((RegistroHorario.Fecha)=Date()))"
        Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
        Do While Not rst.EOF
            turnos = turnos & IIf(ind = 0, Chr(13) & Chr(10), "") & rst("Nombre") & "  " & rst("Apellido") & " --> " & rst("Ingreso")
            rst.MoveNext
            ind = 0
        Loop
        rst.Close
        
        Call EnviarMensajeTelegram(turnos, grupotoperaciones())
    
    Case "etiquetas"
        
        Dim capacidadimpresion As String
        Dim nombreimpresion As String
        Dim cantidadxpaquete As String
                
        Dim produ As String
        Dim counti As Integer
        produ = ""
        
        For counti = 2 To UBound(Detalle) - LBound(Detalle)
            produ = produ & " " & Detalle(counti)
        Next counti
        
        
        Select Case produ
        
            Case "C+C8"
            
                capacidadimpresion = "30+30"
                nombreimpresion = "8gr+8gr"
                cantidadxpaquete = "Cocoa+Cacao 8gr"
                
            Case "C+C10"
            
                capacidadimpresion = "30+30"
                nombreimpresion = "10gr+10gr"
                cantidadxpaquete = "Cocoa+Cacao 10gr"
                
            Case "W+C"
            
                capacidadimpresion = "10+10"
                nombreimpresion = "60gr+2gr"
                cantidadxpaquete = "Cocoa2gr+Waffle60gr"
                
            Case Else
            
                Dim rst1 As DAO.Recordset
                Dim miSQL1 As String

                miSQL1 = "SELECT Cotizaciones.Marca, Cotizaciones.Prioridad, DBIngredientes.Nombre," & _
                " Cotizaciones.Capacidad, Cotizaciones.PaquetePorciones" & _
                " FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
                " WHERE (((Cotizaciones.Marca) = 'Almacen Global') And ((Cotizaciones.Prioridad) = True)" & _
                " And ((DBIngredientes.Nombre) = '" & Detalle(1) & "'))"
                Set rst1 = CurrentDb.OpenRecordset(miSQL1, dbOpenDynaset)
        
                capacidadimpresion = CStr(rst1("Capacidad"))
                nombreimpresion = rst1("Nombre") & " " & CStr(rst1("Capacidad"))
                cantidadxpaquete = CStr(rst1("PaquetePorciones"))
                
                rst1.Close
                
        End Select
        
        Call configurarmargentamañoreporte("PaquetesPorcionesLibre")
        
        'Imprimir
        Dim csti As Integer
        csti = Detalle(1)
        
        DoCmd.OpenReport "PaquetesPorcionesLibre", acViewReport
        [Reports]![PaquetesPorcionesLibre]![apeso] = cantidadxpaquete
        [Reports]![PaquetesPorcionesLibre]![Texto233] = capacidadimpresion
        [Reports]![PaquetesPorcionesLibre]![Texto235] = nombreimpresion
        
        Dim cont As Integer
        cont = 1
        
        Do While cont <= csti
            DoCmd.SelectObject acReport, "PaquetesPorcionesLibre"
            DoCmd.PrintOut acSelection
            cont = cont + 1
        Loop
        DoCmd.Close acReport, "PaquetesPorcionesLibre"

        Call EnviarMensajeTelegram("Impreso", grupotoperaciones())
    Case Else
        Call EnviarMensajeTelegram("No existe formula, volver a intentar", grupotgerencia())
End Select
Exit Sub

Nulo:
Call EnviarMensajeTelegram("Error al momento de ejecutar sentencia, volver a intentar", grupotgerencia())
End Sub

Sub LimpiarHistorialTelegram(codi As Long)
On Error Resume Next
Dim objRequest As Object 'Con lo que se crea la solicitud de internet
Dim datos_posteo As String 'Lo que enviará por mensaje

Dim token As String
token = DLookup("[IdBotTe]", "DatosSistema")


datos_posteo = "offset=" & codi + 1 'Se le muestra al robot que enviar y a que chat
Set objRequest = CreateObject("MSXML2.XMLHTTP") 'Crea un request como archivo XHLM

With objRequest
    .Open "POST", "https://api.telegram.org/bot" & token & "/getUpdates?", False 'Aqui esta la dirección del sitio web con el api del robot
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 'No se que sea
    .send (datos_posteo) 'La indicación de enviar el texto al chat
End With

'MsgBox "historial borrado"
End Sub

Sub EnviarMensajeTelegram(msg As String, grupox As String)
On Error Resume Next
Dim objRequest As Object 'Con lo que se crea la solicitud de internet
Dim datos_posteo As String 'Lo que enviará por mensaje

Dim token, ChatID, Mensaje As String

token = DLookup("[IdBotTe]", "DatosSistema")
ChatID = grupox
Mensaje = msg

datos_posteo = "chat_id=" & ChatID & "&text=" & Mensaje 'Se 'Se le muestra al robot que enviar y a que chat


Set objRequest = CreateObject("MSXML2.XMLHTTP") 'Crea un request como archivo XHLM

With objRequest
    .Open "POST", "https://api.telegram.org/bot" & token & "/sendMessage?", False 'Aqui esta la dirección del sitio web con el api del robot
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 'No se que sea
    .send (datos_posteo) 'La indicación de enviar el texto al chat
End With
End Sub

Sub EnviarImagenTelegram(ruta As String, archivo As String, captio As String, GRUPO As String)
On Error Resume Next

Dim token As String
Dim url As String
Dim METHOD_NAME As String
Dim CHAT_ID As String

If GRUPO = grupotcontabilidad() Then
    token = "6364152209:AAEg0HA59HF9nbiDfKbanKjhG7YxVjfqvWM" ' Token de bot compras
Else
    token = DLookup("[IdBotTe]", "DatosSistema") ' token de bot de sucursal
End If

url = "https://api.telegram.org/bot"
METHOD_NAME = "/sendPhoto?caption=" & captio & "&"
CHAT_ID = GRUPO

Dim FOLDER, JPG_FILE As String
FOLDER = ruta '"C:\Users\LAPTOP\Downloads\"
JPG_FILE = archivo  '"imagen.jpeg"

Dim Data As Object, key
Set Data = CreateObject("Scripting.Dictionary")
Data.Add "chat_id", CHAT_ID

' generate boundary
Dim BOUNDARY, s As String, n As Integer
For n = 1 To 16: s = s & Chr(65 + Int(Rnd * 25)): Next
BOUNDARY = s & CDbl(Now)

Dim part As String, ado As Object
For Each key In Data.Keys
    part = part & "--" & BOUNDARY & vbCrLf
    part = part & "Content-Disposition: form-data; name=""" & key & """" & vbCrLf & vbCrLf
    part = part & Data(key) & vbCrLf
Next
' filename
part = part & "--" & BOUNDARY & vbCrLf
part = part & "Content-Disposition: form-data; name=""photo""; filename=""" & JPG_FILE & """" & vbCrLf & vbCrLf

' read jpg file as binary
Dim jpg
Set ado = CreateObject("ADODB.Stream")
ado.Type = 1 'binary
ado.Open
ado.LoadFromFile FOLDER & JPG_FILE
ado.Position = 0
jpg = ado.Read
ado.Close

' combine part, jpg , end
ado.Open
ado.Position = 0
ado.Type = 1 ' binary
ado.Write ToBytes(part)
ado.Write jpg
ado.Write ToBytes(vbCrLf & "--" & BOUNDARY & "---")
ado.Position = 0

Dim req As Object, reqURL As String
Set req = CreateObject("MSXML2.XMLHTTP")
reqURL = url & token & METHOD_NAME
With req
    .Open "POST", reqURL, False
    .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
    .send ado.Read
    'MsgBox .responseText
End With

End Sub


Function ToBytes(str As String) As Variant

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Open
    ado.Type = 2 ' text
    ado.Charset = "_autodetect"
    ado.WriteText str
    ado.Position = 0
    ado.Type = 1
    ToBytes = ado.Read
    ado.Close
End Function

Sub FotoTemporalTelegram()
On Error Resume Next
Dim carpet As String
Dim arch As String

carpet = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\"
arch = "test.png"
Call TestCamaraPrincipal(carpet & arch)
Call EnviarImagenTelegram(carpet, arch, nombrelocal(), grupotoperaciones())
Kill (carpet & arch)

End Sub



Function grupotpedidospitayacentral() As String
grupotpedidospitayacentral = DLookup("[TelegramPedidosCentral]", "[SistemaGlobal]", 1 = 1)
End Function

Function grupotoperaciones() As String
grupotoperaciones = DLookup("[TelegramOperaciones]", "[SistemaGlobal]", 1 = 1)
End Function

Function grupotgerencia() As String
grupotgerencia = DLookup("[TelegramGerencia]", "[SistemaGlobal]", 1 = 1)
End Function

Function grupotanulaciones() As String
grupotanulaciones = DLookup("[TelegramAnulaciones]", "[SistemaGlobal]", 1 = 1)
End Function

Function grupotcontabilidad() As String
grupotcontabilidad = DLookup("[TelegramContabilidad]", "[SistemaGlobal]", 1 = 1)
End Function

Function detalletextopedido(pedix As Long) As String

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer

miSQL = "SELECT SubPedido.CodSubPedido, SubPedido.CodPedido," & _
" [SubPedido]![Cantidad] & ' ' & nombreproductoventa([SubPedido]![CodBatido]) AS produ," & _
" SubPedido.Vinculo" & _
" FROM SubPedido" & _
" WHERE (((SubPedido.CodPedido) = " & pedix & "))" & _
" ORDER BY SubPedido.CodSubPedido"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
detalletextopedido = ""

For I = 1 To cantli

    detalletextopedido = detalletextopedido & "- " & rst("produ") & vbCrLf
    
    rst.MoveNext
Next I

rst.Close

Exit Function

AgainAgain:

detalletextopedido = ""
End Function


