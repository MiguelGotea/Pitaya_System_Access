' ==========================================================
' Modulo  : Delivery
' Tipo    : 1
' Lineas  : 134
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database
Function KeyGoogleMapsAPI() As String
KeyGoogleMapsAPI = DLookup("[GoogleMapsAPI]", "[SistemaGlobal]")
End Function
Function coordenadassucursal(CodLocal As Integer) As String
coordenadassucursal = DLookup("[Latitude]", "[StatusSucursales]", "[CodLocal]=" & CodLocal) & ", " & DLookup("[Longitude]", "[StatusSucursales]", "[CodLocal]=" & CodLocal)
End Function

Function CalcularDistanciaKM(udesde As String, uhasta As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As New XMLHTTP60
    Dim url As String
    Dim apiKey As String
    Dim responseText As String
    
    ' Tu clave API de Google Maps
    apiKey = KeyGoogleMapsAPI()

    'url
    url = "https://maps.googleapis.com/maps/api/directions/json?" & _
          "origin=" & Replace(udesde, " ", "") & _
          "&destination=" & Replace(uhasta, " ", "") & _
          "&key=" & apiKey
          
    ' Hacer la solicitud HTTP
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        responseText = http.responseText
        
        ' DEBUG: Ver la respuesta completa
        Debug.Print "Respuesta completa: " & responseText
        
        ' Buscar la distancia de manera más precisa
        Dim distancePos As Long
        Dim valuePos As Long
        Dim endPos As Long
        Dim distanciaMetros As Double
        Dim distanciaKM As Double
        
        ' Primero buscar "distance"
        distancePos = InStr(responseText, """distance""")
        If distancePos > 0 Then
            ' Luego buscar "value" dentro del objeto distance
            valuePos = InStr(distancePos, responseText, """value"" : ")
            If valuePos > 0 Then
                ' Extraer el valor numérico
                valuePos = valuePos + 10 ' Longitud de """value"" : "
                endPos = InStr(valuePos, responseText, ",")
                

                If endPos > valuePos Then
                    distanciaMetros = val(Mid(responseText, valuePos, endPos - valuePos))
                    
                    ' Convertir a kilómetros
                    distanciaKM = distanciaMetros / 1000
                    
                    CalcularDistanciaKM = Format(distanciaKM, "0.00") & " km"
                Else
                    CalcularDistanciaKM = "Error: No se pudo extraer valor de distancia"
                End If
            Else
                CalcularDistanciaKM = "Error: No se encontró 'value' en la respuesta"
            End If
        Else
            CalcularDistanciaKM = "Error: No se encontró 'distance' en la respuesta"
        End If
    Else
        CalcularDistanciaKM = "Error en la conexión: " & http.Status
    End If
    
    Exit Function
    
ErrorHandler:
    CalcularDistanciaKM = "Error: " & Err.Description
End Function

Function GenerarURLGoogleMapsCompleto(udesde As String, uhasta As String, Optional modo As String = "optimal") As String
    On Error GoTo ErrorHandler
    
    ' Construir URL más completa
    Dim url As String
    url = "https://www.google.com/maps/dir/" & _
          UrlEncodeGoogleMaps(Replace(Trim(udesde), " ", "")) & "/" & _
          UrlEncodeGoogleMaps(Replace(Trim(uhasta), " ", ""))
    
    ' Agregar parámetros según el modo de transporte
    Select Case LCase(modo)
        Case "optimal"
            url = url & "/@" & ObtenerPuntoMedio(Replace(Trim(udesde), " ", ""), Replace(Trim(uhasta), " ", "")) & ",15z"
        Case "walking", "caminando", "walk"
            url = url & "/data=!3m1!4b1!4m2!4m1!3e2"
        Case "bicycling", "bicicleta", "bike"
            url = url & "/data=!3m1!4b1!4m2!4m1!3e1"
        Case "transit", "transporte", "bus"
            url = url & "/data=!3m1!4b1!4m2!4m1!3e3"
        Case "driving", "conduciendo"
            url = url & "/data=!3m1!4b1!4m2!4m1!3e0"
        Case Else ' "driving", "auto", "car" - MODO CONDUCIR
            url = url & "/data=!3m1!4b1!4m2!4m1!3e0"
    End Select
    
    GenerarURLGoogleMapsCompleto = url
    MsgBox url
    Exit Function
    
ErrorHandler:
    GenerarURLGoogleMapsCompleto = "Error: " & Err.Description

End Function

Function UrlEncodeGoogleMaps(texto As String) As String
    ' Codificar caracteres especiales para URL
    UrlEncodeGoogleMaps = Replace(texto, " ", "+")
    UrlEncodeGoogleMaps = Replace(UrlEncodeGoogleMaps, ",", "%2C")
End Function

Function ObtenerPuntoMedio(coord1 As String, coord2 As String) As String
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double
    
    lat1 = val(Split(coord1, ",")(0))
    lon1 = val(Split(coord1, ",")(1))
    lat2 = val(Split(coord2, ",")(0))
    lon2 = val(Split(coord2, ",")(1))
    
    Dim latMedio As Double, lonMedio As Double
    latMedio = (lat1 + lat2) / 2
    lonMedio = (lon1 + lon2) / 2
    
    ObtenerPuntoMedio = Format(latMedio, "0.000000") & "," & Format(lonMedio, "0.000000")
End Function
