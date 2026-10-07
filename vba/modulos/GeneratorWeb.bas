' ==========================================================
' Modulo  : GeneratorWeb
' Tipo    : 1
' Lineas  : 302
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database
Sub EnviarMensajeWhatsapp(men As String, numbe As String)

On Error GoTo Nulo

CreateObject("Shell.Application").Open "https://wa.me/+505" & numbe & "/?text=" & men & "" 'abrir wsp
Sleep 5000
SendKeys "{ENTER}" 'enviar imagen
Sleep 1000
    
Exit Sub
Nulo:
MsgBox "No se envio mensaje por algun prolema de datos"
End Sub

Sub EnviarImagenWhatsapp(men As String, archi As String, numbe As String)

On Error GoTo Nulo

Application.FollowHyperlink archi ' abrir programa de imagen
Sleep 5000
SendKeys "^c" 'copiar imagen
'Sleep 1000
'Call CreateObject("WScript.Shell").Run("taskkill /f /im " & "PhotosApp.exe", 0, True) 'Cerrar programa de photo
'Call CreateObject("WScript.Shell").Run("taskkill /f /im " & "SnippingTool.exe", 0, True) 'Cerrar programa de corte pantalla
Sleep 3000
CreateObject("Shell.Application").Open "https://wa.me/+505" & numbe & "/?text=" & men & "" 'abrir wsp
CreateObject("Shell.Application").Open "https://wa.me/+505" & numbe & "/?text=" & men & "" 'abrir wsp
Sleep 10000
SendKeys "."
Sleep 3000
SendKeys "^v" 'pegar imagen
Sleep 3000
SendKeys "{ENTER}" 'enviar imagen
Sleep 1000
SendKeys "{NUMLOCK}"
Sleep 1000

Exit Sub
Nulo:
MsgBox "No se envio mensaje por algun prolema de datos"
End Sub

Sub EnviarArchivoWhatsapp(men As String, archi As String, numbe As String)
On Error GoTo Nulo

' Abrir WhatsApp Web con el enlace y esperar
CreateObject("Shell.Application").Open "https://wa.me/+505" & numbe & "/?text=" & men
Sleep 35000

' Cambiar a la ventana del explorador de archivos
SendKeys "+{TAB}"
Sleep 5000
SendKeys "{ENTER}"
Sleep 5000

' Navegar hacia arriba en el explorador de archivos
For I = 1 To 6
    SendKeys "{UP}"
    Sleep 500
Next I

' Abrir la carpeta con el archivo
SendKeys "{ENTER}"
Sleep 10000

' Pegar la dirección del archivo
SendKeys archi
Sleep 2000

' Abrir el archivo
SendKeys "{ENTER}"
Sleep 8000

' Enviar el archivo junto al mensaje
SendKeys "{ENTER}"
Sleep 1000

Exit Sub

Nulo:
    MsgBox "No se envió mensaje por algún problema de datos"
End Sub

Sub GetQRCode(anombre As String, width As Integer, height As Integer)
Dim filepath As String
Dim HttpReq As String
Dim EncContent As String

filepath = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\QR-Insumos\" & anombre & ".png"
EncContent = EncodeURL(anombre)
HttpReq = "http://bwipjs-api.metafloor.com/?bcid=code128&text=" & EncContent
'HttpReq = "https://api.qrserver.com/v1/create-qr-code/?data=" & EncContent & "&size=" & Width & "x" & Height & ""

Call ExportarImagen(HttpReq, filepath)
End Sub

Sub TestCamaraPrincipal(rut As String)

Dim filepath As String
Dim HttpReq As String
Dim canalc, ipc As String
Dim clavecam As String

canalc = DLookup("[CanalCaja]", "[DatosSistema]", 1 = 1)
ipc = DLookup("[IpCamara]", "[DatosSistema]", 1 = 1)
clavecam = DLookup("[ClaveCamara]", "[DatosSistema]", 1 = 1)

filepath = rut
HttpReq = "http://admin:" & clavecam & "@" & ipc & "/ISAPI/Streaming/channels/" & canalc & "/picture"

Call ExportarImagen(HttpReq, filepath)
End Sub

Sub FotoOperarioCheck(Ope As Integer, tipomar As Integer)
Dim filepath As String
Dim HttpReq As String
Dim Direccion1, Direccion2, Direccion3, Direccion4 As String
Dim canalc, ipc As String
Dim Dia, Mes, ano As Long
Dim clavecam As String

Dia = Day(Date)
Mes = Month(Date)
ano = Year(Date)

Direccion1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Check Operarios\" & nombrelocal() & "\" & ano & "\" & Mes & "\" & Dia
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Check Operarios\" & nombrelocal() & "\" & ano & "\" & Mes
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Check Operarios\" & nombrelocal() & "\" & ano
Direccion4 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Check Operarios\" & nombrelocal()

If Dir(Direccion4, vbDirectory) = "" Then
    MkDir Direccion4
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion1, vbDirectory) = "" Then
    MkDir Direccion1
End If

canalc = DLookup("[CanalCaja]", "[DatosSistema]", 1 = 1)
ipc = DLookup("[IpCamara]", "[DatosSistema]", 1 = 1)
clavecam = DLookup("[ClaveCamara]", "[DatosSistema]", 1 = 1)

filepath = Direccion1 & "\" & NombreOperario(Ope) & "_" & IIf(tipomar = 1, "Ingreso", "Salida") & "_" & Hour(Now) & "_" & Minute(Now) & "_" & IIf(Hour(Now) > 12, "PM", "AM") & ".png"
HttpReq = "http://admin:" & clavecam & "@" & ipc & "/ISAPI/Streaming/channels/" & canalc & "/picture"

Call ExportarImagen(HttpReq, filepath)
End Sub

Sub FotoClientePedido(ped As Long)
Dim filepath As String
Dim HttpReq As String
Dim Direccion1, Direccion2, Direccion3, Direccion4 As String
Dim canalc, ipc As String
Dim Dia, Mes, ano As Long

Dia = Day(Date)
Mes = Month(Date)
ano = Year(Date)

Direccion1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\SnapShotClientes\" & nombrelocal() & "\" & ano & "\" & Mes & "\" & Dia
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\SnapShotClientes\" & nombrelocal() & "\" & ano & "\" & Mes
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\SnapShotClientes\" & nombrelocal() & "\" & ano
Direccion4 = "C:\Users\" & NombreSistema() & "\Google Drive BP\SnapShotClientes\" & nombrelocal()

If Dir(Direccion4, vbDirectory) = "" Then
    MkDir Direccion4
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion1, vbDirectory) = "" Then
    MkDir Direccion1
End If

canalc = DLookup("[CanalCaja]", "[DatosSistema]", 1 = 1)
ipc = DLookup("[IpCamara]", "[DatosSistema]", 1 = 1)
Dim clavecam As String
clavecam = DLookup("[ClaveCamara]", "[DatosSistema]", 1 = 1)

filepath = Direccion1 & "\" & ped & ".png"
HttpReq = "http://admin:" & clavecam & "@" & ipc & "/ISAPI/Streaming/channels/" & canalc & "/picture"

Call ExportarImagen(HttpReq, filepath)
End Sub

Sub ExportarImagen(web As String, ruta As String)
On Error GoTo NoSave
Dim ByteData() As Byte
Dim xmlhttp As Object
Dim ReturnContent As String
Dim EncContent As String

Set xmlhttp = CreateObject("MSXML2.XmlHttp")
xmlhttp.Open "GET", web, False
xmlhttp.send
ByteData = xmlhttp.responseBody
Set xmlhttp = Nothing
ReturnContent = StrConv(ByteData, vbUnicode)

Open ruta For Binary As #1
   Put #1, 1, ReturnContent
Close #1

Exit Sub
NoSave:
'MsgBox "No se puede guardar el archivo debido a lo siguiente: " & Err.Description, vbCritical, "File Save Error"
End Sub

Function EncodeURL(str As String) As String
Dim temp As String

temp = Replace(str, " ", "%20")
temp = Replace(temp, "#", "%23")
    
EncodeURL = temp
End Function

Function LeerTxtSolicitud() As String

Dim fieldName As String
fieldName = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\solicitud.txt"
Dim strLineInput As String
Dim fileNum As Integer

fileNum = FreeFile()
Open fieldName For Input As #fileNum
Input #fileNum, strLineInput
LeerTxtSolicitud = strLineInput
Close #fileNum

End Function

Sub EscribirTxtSolicitud(pedi As String)
Dim fieldName As String
fieldName = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\solicitud.txt"
Dim FileNumber As Integer

FileNumber = FreeFile
Open fieldName For Output As FileNumber
Print #FileNumber, pedi
Close FileNumber
End Sub

Function EstaConectadoInternet() As Boolean
Dim objHTTP As Object

'Test for Internet Connection
On Error Resume Next
  Set objHTTP = CreateObject("MSXML2.XMLHTTP")
  objHTTP.Open "GET", "https://www.google.com", False
  objHTTP.send
  IsInternetConnected = (objHTTP.Status = 200)
On Error GoTo 0

EstaConectadoInternet = IsInternetConnected
End Function




Function ObtenerIPv6Publica() As String
On Error GoTo ErrHandler
Dim http As Object
Dim IP As String

Set http = CreateObject("MSXML2.XMLHTTP")
' Servicio que devuelve IP v6 si estás conectado con IPv6
http.Open "GET", "https://api64.ipify.org", False
http.send

IP = http.responseText
ObtenerIPv6Publica = IP

Exit Function

ErrHandler:
ObtenerIPv6Publica = "Error: " & Err.Description
End Function



