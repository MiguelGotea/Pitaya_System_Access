' ==========================================================
' Modulo  : AccessHostinger
' Tipo    : 1  |  Lineas: 457
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:09
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





' ==========================================================
' Exportar una tabla completa a MySQL Hostinger de forma ultra-rápida
' mediante INSERTs masivos por lotes (Batch / Multi-row INSERT)
' ==========================================================
Sub ExportarTablaCompletaHostinger(tablaOrigen As String, tablaDestino As String, Optional Tipooperacion As Integer = 2)
    ' Tipooperacion:
    ' 1 = Anexar
    ' 2 = Vaciar (TRUNCATE/DELETE) y cargar completa
    On Error GoTo ErrorHandler
    
    Dim cn As Object
    Dim rs As DAO.Recordset
    Dim sqlCampos As String
    Dim sqlFila As String
    Dim contador As Long
    Dim totalRegistros As Long
    Dim tamLote As Integer
    Dim I As Integer
    Dim tablaDestinoSQL As String
    
    tamLote = 300 ' Registros por cada viaje de red (óptimo para MySQL en WAN)
    tablaDestinoSQL = "`" & tablaDestino & "`"  ' Backticks para nombres con espacios o palabras reservadas
    
    ' 1. Abrir datos locales en Access (Lectura rápida en memoria)
    Set rs = CurrentDb.OpenRecordset(tablaOrigen, dbOpenSnapshot)
    If rs.EOF And rs.BOF Then
        rs.Close: Set rs = Nothing
        MsgBox "La tabla origen '" & tablaOrigen & "' no contiene registros.", vbInformation
        Exit Sub
    End If
    
    ' 2. Conectar a MySQL Hostinger (charset=utf8mb4 para texto con emojis/Unicode de 4 bytes)
    Set cn = CreateObject("ADODB.Connection")
    cn.Open stringdeconexion() & "charset=utf8mb4;"
    
    ' 3. Limpiar tabla si es Tipooperacion = 2
    If Tipooperacion = 2 Then
        On Error Resume Next
        cn.Execute "TRUNCATE TABLE " & tablaDestinoSQL & ";"
        If Err.Number <> 0 Then
            ' Si TRUNCATE falla por claves foráneas, usamos DELETE
            Err.Clear
            cn.Execute "DELETE FROM " & tablaDestinoSQL & ";"
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' 4. Armar la lista de nombres de campos: (`col1`, `col2`, ...)
    sqlCampos = "INSERT INTO " & tablaDestinoSQL & " ("
    For I = 0 To rs.Fields.Count - 1
        sqlCampos = sqlCampos & "`" & rs.Fields(I).Name & "`"
        If I < rs.Fields.Count - 1 Then sqlCampos = sqlCampos & ", "
    Next I
    sqlCampos = sqlCampos & ") VALUES "
    
    ' 5. Recorrer registros y agrupar en lotes
    contador = 0
    totalRegistros = 0
    
    ' Log de errores por fila (nunca se pierde un registro silenciosamente)
    Dim erroresLog As String
    Dim totalErrores As Long
    erroresLog = ""
    totalErrores = 0
    
    ' Guardar filas del lote actual para retry en caso de fallo del batch
    Dim filasLote() As String
    ReDim filasLote(tamLote - 1)
    
    Do While Not rs.EOF
        ' Construir la tupla: (val1, val2, ...)
        sqlFila = "("
        For I = 0 To rs.Fields.Count - 1
            sqlFila = sqlFila & FormatearValorParaMySQL(rs.Fields(I))
            If I < rs.Fields.Count - 1 Then sqlFila = sqlFila & ", "
        Next I
        sqlFila = sqlFila & ")"
        
        ' Guardar fila en el array del lote para posible retry
        filasLote(contador) = sqlFila
        
        contador = contador + 1
        totalRegistros = totalRegistros + 1
        
        ' Si completamos el lote, intentar enviar; si falla, retry fila a fila
        If contador >= tamLote Then
            Call EnviarLoteConRetry(cn, sqlCampos, filasLote, contador, erroresLog, totalErrores)
            contador = 0
            ReDim filasLote(tamLote - 1)
        End If
        
        rs.MoveNext
    Loop
    
    ' Enviar los registros restantes del último lote si quedaron pendientes
    If contador > 0 Then
        Call EnviarLoteConRetry(cn, sqlCampos, filasLote, contador, erroresLog, totalErrores)
    End If
    
    ' Limpieza
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
    ' Reporte final
    Dim msg As String
    msg = "Tabla '" & tablaOrigen & "' exportada a Hostinger." & vbCrLf & _
          "  Registros procesados : " & totalRegistros & vbCrLf & _
          "  Enviados con \u00e9xito   : " & (totalRegistros - totalErrores) & vbCrLf & _
          "  Errores de fila      : " & totalErrores
    
    If totalErrores > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Detalle de errores:" & vbCrLf & erroresLog
        MsgBox msg, vbExclamation, "Exportaci\u00f3n completada con errores"
    Else
        MsgBox msg, vbInformation, "Exportaci\u00f3n exitosa"
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not cn Is Nothing Then cn.Close: Set cn = Nothing
    MsgBox "Error al exportar tabla '" & tablaOrigen & "': " & Err.Description, vbCritical
End Sub

' --------------------------------------------------------------
' Intenta enviar un lote completo; si falla, reintenta fila a fila
' para nunca perder registros y detectar cu\u00e1l fila espec\u00edfica falla
' --------------------------------------------------------------
Private Sub EnviarLoteConRetry(cn As Object, sqlCampos As String, _
                                filasLote() As String, numFilas As Long, _
                                ByRef erroresLog As String, ByRef totalErrores As Long)
    Dim sqlBatchR As String
    Dim J As Long
    
    ' Armar batch completo
    sqlBatchR = sqlCampos & vbCrLf & filasLote(0)
    For J = 1 To numFilas - 1
        sqlBatchR = sqlBatchR & ", " & vbCrLf & filasLote(J)
    Next J
    
    ' Intentar batch (r\u00e1pido)
    On Error Resume Next
    cn.Execute sqlBatchR & ";"
    Dim batchOK As Boolean
    batchOK = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    
    If batchOK Then Exit Sub  ' \u00c9xito, nada m\u00e1s que hacer
    
    ' El batch fall\u00f3 \u2192 reintentar fila por fila para no perder ning\u00fan registro
    Dim sqlFilaR As String
    For J = 0 To numFilas - 1
        sqlFilaR = sqlCampos & vbCrLf & filasLote(J) & ";"
        On Error Resume Next
        cn.Execute sqlFilaR
        If Err.Number <> 0 Then
            totalErrores = totalErrores + 1
            erroresLog = erroresLog & _
                "  [Fila " & (J + 1) & " del lote] Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                "  Valores: " & Left(filasLote(J), 200) & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next J
End Sub

' -------------------------------------------------------------
' Función auxiliar para formatear tipos de datos de Access a MySQL
' -------------------------------------------------------------
Private Function FormatearValorParaMySQL(fld As DAO.Field) As String
    If IsNull(fld.Value) Then
        FormatearValorParaMySQL = "NULL"
        Exit Function
    End If
    
    Select Case fld.Type
    
        Case dbBoolean
            FormatearValorParaMySQL = IIf(fld.Value = True, "1", "0")
            
        Case dbByte, dbInteger, dbLong
            FormatearValorParaMySQL = CStr(fld.Value)
        
        Case dbBigInt  ' = 16
            FormatearValorParaMySQL = CStr(fld.Value)
            
        Case dbSingle, dbDouble, dbCurrency, dbDecimal, dbNumeric
            ' FIX: 14 decimales para cubrir toda la precisión de Double; punto fijo evita locale español
            FormatearValorParaMySQL = Format(fld.Value, "0.##############")
            
        Case dbDate, dbTime, dbTimeStamp  ' = 8, 22, 23
            ' Guard: MySQL soporta fechas desde 1000-01-01; Access puede guardar fechas anteriores
            If Year(fld.Value) < 1000 Then
                FormatearValorParaMySQL = "NULL"  ' Fecha fuera de rango MySQL ? NULL
            ElseIf Hour(fld.Value) = 0 And Minute(fld.Value) = 0 And Second(fld.Value) = 0 Then
                FormatearValorParaMySQL = "'" & Format(fld.Value, "yyyy-mm-dd") & "'"
            Else
                FormatearValorParaMySQL = "'" & Format(fld.Value, "yyyy-mm-dd hh:nn:ss") & "'"
            End If
        
        Case dbLongBinary, dbBinary, dbVarBinary  ' = 11, 9, 17 — OLE / binarios
            ' No se pueden serializar como texto plano; se omiten como NULL
            FormatearValorParaMySQL = "NULL"
        
        Case dbGUID  ' = 15
            ' Access retorna {xxxxxxxx-...}, MySQL espera 'xxxxxxxx-...' sin llaves
            Dim g As String
            g = CStr(fld.Value)
            g = Replace(Replace(g, "{", ""), "}", "")
            FormatearValorParaMySQL = "'" & g & "'"
        
        Case Else  ' dbText, dbMemo, dbChar y cualquier otro tipo texto
            Dim txt As String
            txt = CStr(fld.Value)
            txt = Replace(txt, "\", "\\")        ' backslash
            txt = Replace(txt, "'", "\'")         ' comilla simple
            txt = Replace(txt, Chr(13), "\r")     ' FIX: carriage return (Memo)
            txt = Replace(txt, Chr(10), "\n")     ' FIX: line feed (Memo)
            txt = Replace(txt, Chr(0), "")        ' FIX: null byte — MySQL lo rechaza
            FormatearValorParaMySQL = "'" & txt & "'"
            
    End Select
End Function




