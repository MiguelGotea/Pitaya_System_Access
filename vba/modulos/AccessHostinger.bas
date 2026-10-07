' ==========================================================
' Modulo  : AccessHostinger
' Tipo    : 1  |  Lineas: 221
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database
Function DriverHostinger() As String
DriverHostinger = DLookup("[DriverHostinger]", "[SistemaGlobal]", 1 = 1)
End Function

Function ServidorHostinger() As String
ServidorHostinger = DLookup("[ServidorHostinger]", "[SistemaGlobal]", 1 = 1)
End Function

Function DBPrincipalHostinger() As String
DBPrincipalHostinger = DLookup("[DBPrincipalHostinger]", "[SistemaGlobal]", 1 = 1)
End Function

Function UsuarioPrincipalHostinger() As String
UsuarioPrincipalHostinger = DLookup("[UsuarioPrincipalHostinger]", "[SistemaGlobal]", 1 = 1)
End Function

Function ClavePrincipalHostinger() As String
ClavePrincipalHostinger = DLookup("[ClavePrincipalHostinger]", "[SistemaGlobal]", 1 = 1)
End Function

Function stringdeconexion() As String
stringdeconexion = "Driver={" & DriverHostinger() & "};" & _
                   "Server=" & ServidorHostinger() & ";" & _
                   "Database=" & DBPrincipalHostinger() & ";" & _
                   "UID=" & UsuarioPrincipalHostinger() & ";" & _
                   "PWD=" & ClavePrincipalHostinger() & ";" & _
                   "Port=3306;" & _
                   "Option=3;"
End Function
Sub importartablasweb()
Call DescargarTablaCompleta("Operarios", "Operarios")
Call DescargarTablaCompleta("AsignacionNivelesCargos", "AsignacionNivelesCargos")
Call DescargarTablaCompleta("NivelesCargos", "NivelesCargos")
End Sub


Sub exportarconsultacondicionalhostinger(consultaorigen As String, tablafinal As String, condi As String, Tipooperacion As Integer)
'Tipooperacion 1: anexar, Tipooperacion 2: vaciar y pegar , Tipooperacion 3: crear tabla
Dim strConn As String
Dim cn As Object

' Conexión ODBC al servidor Hostinger
strConn = stringdeconexion()

' --- Crear consulta temporal con el rango de fechas ---
Dim sqlFiltro As String
Dim nombreTmp As String

nombreTmp = "TablaTemporaParaAnexarConsulta"
If existeTabla(nombreTmp) = 1 Then 'existe tabla
   DoCmd.DeleteObject acTable, nombreTmp ' eliminar si ya existe
End If

sqlFiltro = "SELECT * INTO [" & nombreTmp & "] FROM [" & consultaorigen & "] " & condi

DoCmd.SetWarnings False
DoCmd.RunSQL sqlFiltro
DoCmd.SetWarnings True

' Exportar consulta o tabla
DoCmd.SetWarnings False
DoCmd.TransferDatabase acExport, "ODBC Database", "ODBC;" & strConn, acTable, nombreTmp, "tmptablaconsultadesdeaccess"
DoCmd.SetWarnings True

' Crear conexión ADO
Set cn = CreateObject("ADODB.Connection")
cn.Open strConn

Select Case Tipooperacion
    Case 1
        ' Ejecutar el INSERT
        cn.Execute "INSERT INTO " & tablafinal & " SELECT * FROM tmptablaconsultadesdeaccess;"
        
    Case 2
        ' --- Limpiar la tabla destino ---
        cn.Execute "DELETE FROM " & tablafinal & ";"
        
        ' Ejecutar el INSERT
        cn.Execute "INSERT INTO " & tablafinal & " SELECT * FROM tmptablaconsultadesdeaccess;"
        
    Case 3
        ' --- Borrar la tabla destino si existe ---
        On Error Resume Next
        cn.Execute "DROP TABLE " & tablafinal & ";"
        On Error GoTo 0
        
        ' --- Crear tabla nueva con los datos ---
        cn.Execute "CREATE TABLE " & tablafinal & " AS SELECT * FROM tmptablaconsultadesdeaccess;"
    
    Case Else
        Exit Sub
    
End Select

' Ejecutar el DROP
cn.Execute "DROP TABLE tmptablaconsultadesdeaccess;"

' Borra tabla que se creo del filtro
If existeTabla(nombreTmp) = 1 Then 'existe tabla
   DoCmd.DeleteObject acTable, nombreTmp ' eliminar si ya existe
End If

' Cerrar conexión
cn.Close
Set cn = Nothing

MsgBox "Datos exportados con éxito"

End Sub


Sub ExportarCSVConsultaCondicional(consu As String, condi As String, ruta As String)
'si se exporta una tabla no puede contener fechas porque no se le paso por la funcion cfechafechasql
On Error GoTo ErrorHandler

' --- Crear consulta temporal con el rango de fechas ---
Dim sqlFiltro As String
Dim nombreTmp As String

nombreTmp = "TablaTemporaParaAnexarConsulta"
If existeTabla(nombreTmp) = 1 Then 'existe tabla
   DoCmd.DeleteObject acTable, nombreTmp ' eliminar si ya existe
End If

sqlFiltro = "SELECT * INTO [" & nombreTmp & "] FROM [" & consu & "] " & condi

DoCmd.SetWarnings False
DoCmd.RunSQL sqlFiltro
DoCmd.SetWarnings True
    
' Exportar normalmente (ANSI)
DoCmd.SetWarnings False
DoCmd.TransferText TransferType:=acExportDelim, TableName:=nombreTmp, filename:=ruta, HasFieldNames:=True
DoCmd.SetWarnings True

' === Convertir a UTF-8 ===
Dim stmIn As Object, stmOut As Object

Set stmIn = CreateObject("ADODB.Stream")
stmIn.Type = 2
stmIn.Charset = "Windows-1252"   ' lo que generó Access
stmIn.Open
stmIn.LoadFromFile ruta

Set stmOut = CreateObject("ADODB.Stream")
stmOut.Type = 2
stmOut.Charset = "UTF-8"
stmOut.Open
stmIn.CopyTo stmOut
stmOut.SaveToFile ruta, 2   ' 2 = sobrescribir

stmIn.Close
stmOut.Close

' --- Borrar la tabla temporal ---
DoCmd.DeleteObject acTable, nombreTmp

'MsgBox "Consulta exportada exitosamente en UTF-8", vbInformation
Exit Sub

ErrorHandler:
MsgBox "Error al exportar: " & Err.Description, vbCritical
End Sub

Sub ExportarCSVTablaEspecifica(tablac As String, ruta As String)
'si se exporta una tabla no puede contener fechas porque no se le paso por la funcion cfechafechasql
On Error GoTo ErrorHandler
    
' Exportar normalmente (ANSI)
DoCmd.SetWarnings False
DoCmd.TransferText TransferType:=acExportDelim, TableName:=tablac, filename:=ruta, HasFieldNames:=True
DoCmd.SetWarnings True

' === Convertir a UTF-8 ===
Dim stmIn As Object, stmOut As Object

Set stmIn = CreateObject("ADODB.Stream")
stmIn.Type = 2
stmIn.Charset = "Windows-1252"   ' lo que generó Access
stmIn.Open
stmIn.LoadFromFile ruta

Set stmOut = CreateObject("ADODB.Stream")
stmOut.Type = 2
stmOut.Charset = "UTF-8"
stmOut.Open
stmIn.CopyTo stmOut
stmOut.SaveToFile ruta, 2   ' 2 = sobrescribir

stmIn.Close
stmOut.Close

'MsgBox "Consulta exportada exitosamente en UTF-8", vbInformation
Exit Sub

ErrorHandler:
MsgBox "Error al exportar: " & Err.Description, vbCritical
End Sub


Sub ActualizarArchivoCSVVentasFecha(fechi As Date)
Dim ruta As String
Dim condic As String
ruta = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Web\Actualizacion Ventas Automatico\" & codigoLocal() & "_" & cfechasqlfecha(fechi) & ".csv"
condic = "WHERE Fecha BETWEEN '" & cfechasqlfecha(fechi) & "' AND '" & cfechasqlfecha(fechi) & "'"
Call ExportarCSVConsultaCondicional("ResumenVentasMesExcelHostingerCSV", condic, ruta)
End Sub


Sub ActualizarArchivoCSVMembresias()
Dim ruta As String
Dim condic As String


'Trigge de Apps script se ejecuta entre las 12 y la 1, en ese intervalo mejor no crear nada

ruta = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Web\Actualizacion Membresias Automatico\" & codigoLocal() & ".csv"
condic = ""
Call ExportarCSVConsultaCondicional("ResumenMembresiasHostingerCSV", condic, ruta)
End Sub
