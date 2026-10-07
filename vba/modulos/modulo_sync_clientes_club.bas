' ==========================================================
' Modulo  : modulo_sync_clientes_club
' Tipo    : 1
' Lineas  : 285
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:31
' ==========================================================
Option Compare Database

' ========== SUB: SINCRONIZAR CÉDULAS LOCALES ==========
' Descarga y actualiza las cédulas del host para la sucursal actual
Public Sub SincronizarCedulasLocales()
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim Sucursal As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalActualizados As Long
    Dim totalRecibidos As Long
    Const API_BASE_URL = "https://proxy.batidospitaya.com/api/MS_Access/clientes_club/"
    Const API_TOKEN = "a8f5e2d9c4b7a1e6f3d8c5b2a9e6d3f0c7a4b1e8d5c2a9f6e3d0c7b4a1e8f5d2"
    
    ' 1. Obtener sucursal local (función del sistema)
    Sucursal = codigoLocal()
    
    If IsNull(Sucursal) Or Sucursal = "" Then
        MsgBox "No se pudo determinar el código de sucursal local.", vbExclamation, "Error de Sincronización"
        Exit Sub
    End If
    
    ' 2. Llamar API
    url = API_BASE_URL & "obtener_cedulas_sucursal.php?token=" & API_TOKEN & "&sucursal=" & Sucursal
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    
    response = http.responseText
    
    ' 3. Procesar respuesta
    If InStr(response, """success"":true") = 0 Then
        MsgBox "Error en la respuesta del servidor: " & response, vbCritical, "Error de API"
        GoTo Cleanup
    End If
    
    ' 4. Parsear JSON simple (Array de objetos)
    ' Ejemplo: {"datos":[{"membresia":123,"cedula":"123456"},{"membresia":456,"cedula":"789012"}]}
    
    totalActualizados = 0
    totalRecibidos = 0
    
    Set db = CurrentDb
    
    ' Extraer la parte de "datos":[...]
    Dim datosPart As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(response, """datos"":[") + 8
    endPos = InStrRev(response, "]")
    
    If startPos > 8 And endPos > startPos Then
        datosPart = Mid(response, startPos, endPos - startPos)
        
        ' Dividir por objetos {}
        Dim objetos() As String
        objetos = Split(datosPart, "},{")
        
        Dim obj As Variant
        Dim membresia As String
        Dim Cedula As String
        
        For Each obj In objetos
            ' Limpiar llaves si existen
            Dim cleanObj As String
            cleanObj = Replace(Replace(obj, "{", ""), "}", "")
            
            ' Extraer membresia
            membresia = ParseSimpleJSONValue(cleanObj, "membresia")
            ' Extraer cedula
            Cedula = ParseSimpleJSONValue(cleanObj, "cedula")
            
            If membresia <> "" And Cedula <> "" Then
                totalRecibidos = totalRecibidos + 1
                
                ' Actualizar tabla local ClientesClub
                ' Usamos CodCliente (Access) que corresponde a membresia (Host)
                db.Execute "UPDATE ClientesClub SET Cedula = '" & Replace(Cedula, "'", "''") & "' " & _
                           "WHERE CodCliente = " & membresia, dbFailOnError
                
                If db.RecordsAffected > 0 Then
                    totalActualizados = totalActualizados + 1
                End If
            End If
        Next obj
    End If
    
    MsgBox "Sincronización completada." & vbCrLf & _
           "Sucursal: " & Sucursal & vbCrLf & _
           "Registros recibidos: " & totalRecibidos & vbCrLf & _
           "Localmente actualizados: " & totalActualizados, vbInformation, "Éxito"

Cleanup:
    Set http = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error durante la sincronización: " & Err.Description, vbCritical, "Error Crítico"
    Resume Cleanup
End Sub


' ========== SUB: SINCRONIZAR DATOS COMPLETOS LOCALES ==========
' Descarga y actualiza desde el host: Cedula, Nombre, Apellidos, Celular, Cumpleanos y Correo
' para todos los clientes de la sucursal actual.
' Solo sobreescribe campos que vengan con valor en el host (no pisa datos locales con nulos).
Public Sub SincronizarDatosClientesLocales()
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim Sucursal As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalActualizados As Long
    Dim totalRecibidos As Long
    Const API_BASE_URL = "https://proxy.batidospitaya.com/api/MS_Access/clientes_club/"
    Const API_TOKEN = "a8f5e2d9c4b7a1e6f3d8c5b2a9e6d3f0c7a4b1e8d5c2a9f6e3d0c7b4a1e8f5d2"

    ' 1. Obtener sucursal local (función del sistema)
    Sucursal = codigoLocal()

    If IsNull(Sucursal) Or Sucursal = "" Then
        MsgBox "No se pudo determinar el código de sucursal local.", vbExclamation, "Error de Sincronización"
        Exit Sub
    End If

    ' 2. Llamar API con el nuevo endpoint de datos completos
    url = API_BASE_URL & "obtener_datos_clientes_sucursal.php?token=" & API_TOKEN & "&sucursal=" & Sucursal

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.send

    response = http.responseText

    ' 3. Verificar respuesta
    If InStr(response, """success"":true") = 0 Then
        MsgBox "Error en la respuesta del servidor: " & response, vbCritical, "Error de API"
        GoTo Cleanup
    End If

    ' 4. Parsear JSON y actualizar tabla local ClientesClub
    ' Estructura esperada: {"datos":[{"membresia":123,"nombre":"...","apellido":"...","cedula":"...","celular":"...","fecha_nacimiento":"YYYY-MM-DD","correo":"..."},...]}

    totalActualizados = 0
    totalRecibidos = 0

    Set db = CurrentDb

    ' Extraer la parte de "datos":[...]
    Dim datosPart As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(response, """datos"":[") + 8
    endPos = InStrRev(response, "]")

    If startPos > 8 And endPos > startPos Then
        datosPart = Mid(response, startPos, endPos - startPos)

        ' Dividir por objetos {}
        Dim objetos() As String
        objetos = Split(datosPart, "},{")

        Dim obj As Variant
        Dim membresia As String
        Dim Nombre As String
        Dim Apellido As String
        Dim Cedula As String
        Dim Celular As String
        Dim fechaNacimiento As String
        Dim Correo As String

        For Each obj In objetos
            ' Limpiar llaves si existen
            Dim cleanObj As String
            cleanObj = Replace(Replace(obj, "{", ""), "}", "")

            ' Extraer todos los campos
            membresia = ParseSimpleJSONValue(cleanObj, "membresia")
            Nombre = ParseSimpleJSONValue(cleanObj, "nombre")
            Apellido = ParseSimpleJSONValue(cleanObj, "apellido")
            Cedula = ParseSimpleJSONValue(cleanObj, "cedula")
            Celular = ParseSimpleJSONValue(cleanObj, "celular")
            fechaNacimiento = ParseSimpleJSONValue(cleanObj, "fecha_nacimiento")
            Correo = ParseSimpleJSONValue(cleanObj, "correo")

            ' Solo procesar si tenemos un código de cliente válido
            If membresia <> "" And membresia <> "null" Then
                totalRecibidos = totalRecibidos + 1

                ' Construir cláusula SET solo con los valores que vienen del host
                ' (si el host envía null/vacío no pisamos el dato local)
                Dim setParts As String
                setParts = ""

                If Nombre <> "" And Nombre <> "null" Then
                    setParts = setParts & "Nombre = '" & Replace(Nombre, "'", "''") & "', "
                End If

                If Apellido <> "" And Apellido <> "null" Then
                    setParts = setParts & "Apellidos = '" & Replace(Apellido, "'", "''") & "', "
                End If

                If Cedula <> "" And Cedula <> "null" Then
                    setParts = setParts & "Cedula = '" & Replace(Cedula, "'", "''") & "', "
                End If

                If Celular <> "" And Celular <> "null" Then
                    setParts = setParts & "Celular = '" & Replace(Celular, "'", "''") & "', "
                End If

                ' Fecha: Access SQL usa formato #YYYY-MM-DD#
                If fechaNacimiento <> "" And fechaNacimiento <> "null" Then
                    setParts = setParts & "Cumpleanos = #" & fechaNacimiento & "#, "
                End If

                If Correo <> "" And Correo <> "null" Then
                    setParts = setParts & "Correo = '" & Replace(Correo, "'", "''") & "', "
                End If

                ' Solo ejecutar UPDATE si hay al menos un campo con valor
                If Len(setParts) > 0 Then
                    ' Quitar coma final antes de WHERE
                    setParts = Left(setParts, Len(setParts) - 2)

                    db.Execute "UPDATE ClientesClub SET " & setParts & _
                               " WHERE CodCliente = " & membresia, dbFailOnError

                    If db.RecordsAffected > 0 Then
                        totalActualizados = totalActualizados + 1
                    End If
                End If
            End If
        Next obj
    End If

    'MsgBox "Sincronización de datos completada." & vbCrLf & _
           "Sucursal: " & Sucursal & vbCrLf & _
           "Registros recibidos: " & totalRecibidos & vbCrLf & _
           "Localmente actualizados: " & totalActualizados, vbInformation, "Éxito"

Cleanup:
    Set http = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error durante la sincronización de datos: " & Err.Description, vbCritical, "Error Crítico"
    Resume Cleanup
End Sub

' Función auxiliar para extraer valores de un objeto JSON simple "key":value o "key":"value"
Public Function ParseSimpleJSONValue(JSONStr As String, key As String) As String
    Dim keySearch As String
    Dim p1 As Long, p2 As Long
    
    keySearch = """" & key & """:"
    p1 = InStr(JSONStr, keySearch)
    
    If p1 > 0 Then
        p1 = p1 + Len(keySearch)
        ' Verificar si el valor empieza con comilla
        If Mid(JSONStr, p1, 1) = """" Then
            p1 = p1 + 1
            p2 = InStr(p1, JSONStr, """")
        Else
            ' Es numérico o null, buscar coma o fin
            p2 = InStr(p1, JSONStr, ",")
            If p2 = 0 Then p2 = Len(JSONStr) + 1
        End If
        
        If p2 > p1 Then
            ParseSimpleJSONValue = Mid(JSONStr, p1, p2 - p1)
        End If
    End If
End Function

