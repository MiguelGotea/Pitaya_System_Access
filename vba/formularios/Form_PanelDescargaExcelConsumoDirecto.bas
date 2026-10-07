' ==========================================================
' Modulo  : Form_PanelDescargaExcelConsumoDirecto
' Tipo    : 100  |  Lineas: 1374
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:20
' ==========================================================

Option Compare Database





Private Sub Comando133_Click()

    ' Declaraciones
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object ' Excel.Application
    Dim xlBook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim strSQL As String
    Dim I As Long, j As Long
    Dim semana_desde As Integer
    Dim semana_hasta As Integer
    Dim numSemanas As Integer
    Dim fila As Long
    Dim semana_actual As Integer
    Dim resultado As Variant
    Dim startTime As Single
    
    ' Declaraciones para el manejo de archivos
    Dim Direccion As String
    Dim archivo As String
    
    startTime = Timer ' Para medir tiempo de ejecución
    
    ' Solicitar rango de semanas
    semana_desde = Me.semana_desde
    semana_hasta = Me.semana_hasta
    
    ' Validar entrada
    If semana_desde <= 0 Or semana_hasta < semana_desde Then
        MsgBox "Rango de semanas inválido. La semana inicial debe ser mayor a 0 y la semana final debe ser mayor o igual a la inicial.", vbExclamation
        Exit Sub
    End If
    
    numSemanas = semana_hasta - semana_desde + 1
    
    ' Mostrar mensaje de progreso
    DoCmd.Hourglass True
    DoCmd.Echo False, "Generando reporte de ventas por producto. Por favor espere..."
    
    ' Crear instancia de Excel INVISIBLE
    On Error GoTo ErrorExcel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' No mostrar Excel
    xlApp.DisplayAlerts = False ' No mostrar alertas
    xlApp.ScreenUpdating = False ' No actualizar pantalla (más rápido)
    
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error GoTo 0
    
    ' Configurar la hoja
    With xlSheet
        .Name = "Ventas Semanas " & semana_desde & "-" & semana_hasta
        
        ' Encabezados
        .Cells(1, 1).Value = "Producto"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.color = RGB(200, 200, 200)
        
        ' Encabezados de semanas
        For I = 1 To numSemanas
            .Cells(1, I + 1).Value = "Semana " & (semana_desde + I - 1)
            .Cells(1, I + 1).Font.Bold = True
            .Cells(1, I + 1).Interior.color = RGB(200, 200, 200)
            .Cells(1, I + 1).horizontalAlignment = -4108 ' xlCenter
        Next I
    End With
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    
    ' Consulta SQL optimizada
    strSQL = "SELECT DISTINCT SubReceta.codporcion, DBIngredientes.Nombre, TiposVariables.Orden " & _
             "FROM ((DBBatidos " & _
             "INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) " & _
             "INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente) " & _
             "INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo " & _
             "WHERE TiposVariables.Control = True " & _
             "AND DBBatidos.Vigencia = True " & _
             "AND SubReceta.codporcion Is Not Null " & _
             "AND (DBBatidos.CodGrupo <> 7 OR DBBatidos.CodGrupo Is Null) " & _
             "AND PorcionDentroDeMezcla(SubReceta.codporcion) = 0 " & _
             "ORDER BY TiposVariables.Orden, DBIngredientes.Nombre;"
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No se encontraron registros para exportar.", vbInformation
        GoTo Limpiar
    End If
    
    ' Posición inicial para datos
    fila = 2
    
    ' Obtener total de registros
    Dim totalRegistros As Long
    rs.MoveLast
    totalRegistros = rs.RecordCount
    rs.MoveFirst
    
    ' Crear arrays para mejorar el rendimiento (opcional, pero más rápido)
    Dim datosProductos() As Variant
    Dim datosVentas() As Variant
    Dim productoActual As String
    
    ' Recorrer el recordset y llenar datos directamente en Excel (sin arrays)
    Do While Not rs.EOF
        ' Actualizar mensaje de progreso en Access
        If (fila - 1) Mod 25 = 0 Then
            DoCmd.Echo False, "Procesando registro " & (fila - 1) & " de " & totalRegistros & "..."
            DoEvents ' Permite procesar otros eventos
        End If
        
        ' Nombre del producto usando la función nombreproductocoti
        On Error Resume Next
        resultado = nombreproductocoti(rs!codporcion)
        If Err.Number <> 0 Then
            xlSheet.Cells(fila, 1).Value = "Error en " & rs!codporcion
            Err.Clear
        Else
            xlSheet.Cells(fila, 1).Value = resultado
        End If
        On Error GoTo 0
        
        ' Calcular ventas para cada semana individualmente
        For j = 1 To numSemanas
            semana_actual = semana_desde + j - 1
            
            On Error Resume Next
            ' Llamar a la función con un solo parámetro de semana
            resultado = VentasXTipoPorcionSemana(semana_actual, rs!codporcion)
            
            If Err.Number <> 0 Then
                xlSheet.Cells(fila, j + 1).Value = 0 ' Poner 0 en caso de error
                Err.Clear
            Else
                If IsNumeric(resultado) Then
                    xlSheet.Cells(fila, j + 1).Value = resultado
                Else
                    xlSheet.Cells(fila, j + 1).Value = 0
                End If
            End If
            On Error GoTo 0
        Next j
        
        fila = fila + 1
        rs.MoveNext
    Loop
    
    ' Aplicar formato a todo el rango de una sola vez (más rápido)
    If fila > 2 Then
        With xlSheet
            ' Autoajustar columna de productos
            .Columns("A").EntireColumn.AutoFit
            
            ' Formato de números para las columnas de ventas
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).NumberFormat = "#,##0.00"
            
            ' Agregar bordes a toda la tabla
            .Range(.Cells(1, 1), .Cells(fila - 1, numSemanas + 1)).Borders.LineStyle = 1 ' xlContinuous
            
            ' Agregar color alternado a las filas para mejor lectura (más rápido aplicarlo por rangos)
            For I = 2 To fila - 1 Step 2
                .Range(.Cells(I, 1), .Cells(I, numSemanas + 1)).Interior.color = RGB(240, 240, 240)
            Next I
            
            ' Centrar los datos de las semanas
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).horizontalAlignment = -4152 ' xlRight
        End With
    End If
    
    ' --- Guardar archivo ---
    
    ' Construir la ruta del directorio
    On Error Resume Next
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
    
    ' Verificar si el directorio existe, si no, crearlo
    If Dir(Direccion, vbDirectory) = "" Then
        MkDir Direccion
    End If
    
    ' Construir el nombre del archivo
    archivo = Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - Sem " & semana_desde & " hasta " & semana_hasta & " - Porciones.xlsx"
    
    ' Guardar el libro
    xlBook.SaveAs archivo, 51 ' 51 = xlOpenXMLWorkbook (formato .xlsx)
    
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo: " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Cerrar Excel sin guardar cambios adicionales
    xlBook.Close False
    xlApp.Quit
    
    ' Restaurar cursor y mensajes
    DoCmd.Hourglass False
    DoCmd.Echo True

Limpiar:
    ' Limpiar objetos
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True
    On Error GoTo 0
    Exit Sub

ErrorExcel:
    MsgBox "Error al crear Excel. Asegúrate de tener Microsoft Excel instalado.", vbCritical
    Resume Limpiar
End Sub


Private Sub Comando332_Click()

    ' Declaraciones
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object ' Excel.Application
    Dim xlBook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim strSQL As String
    Dim I As Long, j As Long
    Dim semana_desde As Integer
    Dim semana_hasta As Integer
    Dim numSemanas As Integer
    Dim fila As Long
    Dim semana_actual As Integer
    Dim resultado As Variant
    Dim Conversion As Double
    Dim startTime As Single
    
    ' Declaraciones para el manejo de archivos
    Dim Direccion As String
    Dim archivo As String
    
    startTime = Timer ' Para medir tiempo de ejecución
    
    ' Solicitar rango de semanas
    semana_desde = Me.semana_desde
    semana_hasta = Me.semana_hasta
    
    
    ' Validar entrada
    If semana_desde <= 0 Or semana_hasta < semana_desde Then
        MsgBox "Rango de semanas inválido. La semana inicial debe ser mayor a 0 y la semana final debe ser mayor o igual a la inicial.", vbExclamation
        Exit Sub
    End If
    
    numSemanas = semana_hasta - semana_desde + 1
    
    ' Mostrar mensaje de progreso
    DoCmd.Hourglass True
    DoCmd.Echo False, "Generando reporte de consumos por producto. Por favor espere..."
    
    ' Crear instancia de Excel INVISIBLE
    On Error GoTo ErrorExcel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error GoTo 0
    
    ' Configurar la hoja
    With xlSheet
        .Name = "Consumos Sem " & semana_desde & "-" & semana_hasta
        
        ' Encabezados
        .Cells(1, 1).Value = "Producto (nombreproductocoti)"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.color = RGB(200, 200, 200)
        .Cells(1, 1).ColumnWidth = 40
        
        ' Encabezados de semanas
        For I = 1 To numSemanas
            .Cells(1, I + 1).Value = "Semana " & (semana_desde + I - 1)
            .Cells(1, I + 1).Font.Bold = True
            .Cells(1, I + 1).Interior.color = RGB(200, 200, 200)
            .Cells(1, I + 1).horizontalAlignment = -4108 ' xlCenter
            .Cells(1, I + 1).ColumnWidth = 15
        Next I
        
        ' Título del reporte
        .Rows("1:1").RowHeight = 20
    End With
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    
    ' Consulta SQL corregida - SIN DISTINCT y con todos los campos necesarios
    strSQL = "SELECT " & _
             "TiposVariables.Orden, " & _
             "DBIngredientes.Nombre, " & _
             "SubReceta.codporcion, " & _
             "CotiPrincipalDeIngrediente(SubReceta.CodIngrediente) AS coti, " & _
             "Cotizaciones.CodCotizacion, " & _
             "DBIngredientes.CodIngrediente, " & _
             "DBIngredientes.Tipo " & _
             "FROM ((DBBatidos " & _
             "INNER JOIN (DBIngredientes " & _
             "INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) " & _
             "ON DBBatidos.CodBatido = SubReceta.CodBatido) " & _
             "INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo) " & _
             "INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente " & _
             "WHERE TiposVariables.Control = True " & _
             "AND DBBatidos.Vigencia = True " & _
             "AND SubReceta.codporcion Is Null " & _
             "AND (DBBatidos.CodGrupo <> 7 OR DBBatidos.CodGrupo Is Null) " & _
             "AND CotisNoPorcionPrioridadVigente(Cotizaciones.CodCotizacion) <>0 " & _
             "GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, SubReceta.codporcion, " & _
             "CotiPrincipalDeIngrediente(SubReceta.CodIngrediente), Cotizaciones.CodCotizacion, " & _
             "DBIngredientes.CodIngrediente, DBIngredientes.Tipo " & _
             "ORDER BY TiposVariables.Orden, DBIngredientes.Nombre, DBIngredientes.Tipo;"
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No se encontraron registros para exportar.", vbInformation
        GoTo Limpiar
    End If
    
    ' Posición inicial para datos
    fila = 2
    
    ' Obtener total de registros
    Dim totalRegistros As Long
    rs.MoveLast
    totalRegistros = rs.RecordCount
    rs.MoveFirst
    
    ' Recorrer el recordset y llenar datos
    Do While Not rs.EOF
        ' Actualizar mensaje de progreso
        If (fila - 1) Mod 25 = 0 Then
            DoCmd.Echo False, "Procesando registro " & (fila - 1) & " de " & totalRegistros & "..."
            DoEvents
        End If
        
        ' Nombre del producto usando nombreproductocoti con CodCotizacion
        On Error Resume Next
        resultado = nombreproductocoti(rs!CodCotizacion)
        If Err.Number <> 0 Then
            xlSheet.Cells(fila, 1).Value = "Error en Cotizacion " & rs!CodCotizacion
            Err.Clear
        Else
            xlSheet.Cells(fila, 1).Value = resultado
        End If
        On Error GoTo 0
        
        ' Obtener el valor de conversión para este producto
        On Error Resume Next
        Conversion = ObtenerConversion(rs!CodCotizacion)
        If Err.Number <> 0 Or Conversion = 0 Then
            Conversion = 1 ' Valor por defecto para evitar división por cero
            Err.Clear
        End If
        On Error GoTo 0
        
        ' Calcular consumos para cada semana
        For j = 1 To numSemanas
            semana_actual = semana_desde + j - 1
            
            On Error Resume Next
            ' Calcular: VentasXTipoNoPorcionSemana(semana, CodIngrediente) / conversion
            resultado = VentasXTipoNoPorcionSemana(semana_actual, rs!CodIngrediente)
            
            If Err.Number <> 0 Then
                xlSheet.Cells(fila, j + 1).Value = 0
                Err.Clear
            Else
                If IsNumeric(resultado) And Conversion <> 0 Then
                    xlSheet.Cells(fila, j + 1).Value = resultado / Conversion
                Else
                    xlSheet.Cells(fila, j + 1).Value = 0
                End If
            End If
            On Error GoTo 0
        Next j
        
        fila = fila + 1
        rs.MoveNext
    Loop
    
    ' Aplicar formato
    If fila > 2 Then
        With xlSheet
            ' Autoajustar columna de productos
            .Columns("A").EntireColumn.AutoFit
            If .Columns("A").ColumnWidth > 80 Then .Columns("A").ColumnWidth = 80
            
            ' Formato de números para las columnas de consumos
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).NumberFormat = "#,##0.00"
            
            ' Agregar bordes
            .Range(.Cells(1, 1), .Cells(fila - 1, numSemanas + 1)).Borders.LineStyle = 1
            
            ' Filas con color alternado
            For I = 2 To fila - 1 Step 2
                .Range(.Cells(I, 1), .Cells(I, numSemanas + 1)).Interior.color = RGB(240, 240, 240)
            Next I
            
            ' Alinear datos a la derecha
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).horizontalAlignment = -4152
        End With
    End If
    
    ' --- Guardar archivo ---
    
    ' Construir la ruta del directorio
    On Error Resume Next
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
    
    ' Verificar si el directorio existe, si no, crearlo
    If Dir(Direccion, vbDirectory) = "" Then
        MkDir Direccion
    End If
    
    ' Construir el nombre del archivo (CORREGIDO según tu especificación)
    archivo = Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - Sem " & semana_desde & " hasta " & semana_hasta & " - No Porciones.xlsx"
    
    ' Guardar el libro
    xlBook.SaveAs archivo, 51
    
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo: " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Cerrar Excel sin guardar cambios adicionales
    xlBook.Close False
    xlApp.Quit

    ' Restaurar cursor y mensajes
    DoCmd.Hourglass False
    DoCmd.Echo True

Limpiar:
    ' Limpiar objetos
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True
    On Error GoTo 0
    Exit Sub

ErrorExcel:
    MsgBox "Error al crear Excel. Asegúrate de tener Microsoft Excel instalado.", vbCritical
    Resume Limpiar
End Sub

' Función auxiliar para obtener el valor de conversión
Function ObtenerConversion(CodCotizacion As Variant) As Double
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    sql = "SELECT [Conversion] FROM [Cotizaciones] WHERE [CodCotizacion] = " & CodCotizacion
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        If Not IsNull(rs!Conversion) Then
            ObtenerConversion = rs!Conversion
        Else
            ObtenerConversion = 0
        End If
    Else
        ObtenerConversion = 0
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ObtenerConversion = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function


Private Sub Comando334_Click()

    ' Declaraciones
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object ' Excel.Application
    Dim xlBook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim strSQL As String
    Dim I As Long, j As Long
    Dim semana_desde As Integer
    Dim semana_hasta As Integer
    Dim numSemanas As Integer
    Dim fila As Long
    Dim semana_actual As Integer
    Dim resultado As Variant
    Dim divisor As Double
    Dim NombreProducto As String
    Dim startTime As Single
    
    ' Declaraciones para el manejo de archivos
    Dim Direccion As String
    Dim archivo As String
    
    startTime = Timer ' Para medir tiempo de ejecución
    
    ' Solicitar rango de semanas
    semana_desde = Me.semana_desde
    semana_hasta = Me.semana_hasta
    
    ' Validar entrada
    If semana_desde <= 0 Or semana_hasta < semana_desde Then
        MsgBox "Rango de semanas inválido. La semana inicial debe ser mayor a 0 y la semana final debe ser mayor o igual a la inicial.", vbExclamation
        Exit Sub
    End If
    
    numSemanas = semana_hasta - semana_desde + 1
    
    ' Mostrar mensaje de progreso
    DoCmd.Hourglass True
    DoCmd.Echo False, "Generando reporte de consumos compacto por producto. Por favor espere..."
    
    ' Crear instancia de Excel INVISIBLE
    On Error GoTo ErrorExcel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error GoTo 0
    
    ' Configurar la hoja
    With xlSheet
        .Name = "Compacto Sem " & semana_desde & "-" & semana_hasta
        
        ' Encabezados
        .Cells(1, 1).Value = "Producto"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.color = RGB(200, 200, 200)
        .Cells(1, 1).ColumnWidth = 40
        
        ' Encabezados de semanas
        For I = 1 To numSemanas
            .Cells(1, I + 1).Value = "Semana " & (semana_desde + I - 1)
            .Cells(1, I + 1).Font.Bold = True
            .Cells(1, I + 1).Interior.color = RGB(200, 200, 200)
            .Cells(1, I + 1).horizontalAlignment = -4108 ' xlCenter
            .Cells(1, I + 1).ColumnWidth = 15
        Next I
        
        ' Título del reporte
        .Rows("1:1").RowHeight = 20
    End With
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    
    ' Consulta SQL para el reporte compacto
    strSQL = "SELECT " & _
             "Left(DBIngredientes.CodIngrediente,1) AS ini, " & _
             "DBIngredientes.Tipo, " & _
             "DBIngredientes.Nombre, " & _
             "SubReceta.codporcion, " & _
             "SubReceta.CodIngrediente, " & _
             "DBIngredientes.Unidad " & _
             "FROM (((DBIngredientes " & _
             "INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) " & _
             "INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) " & _
             "INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) " & _
             "INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo " & _
             "WHERE DBBatidos.Vigencia = True " & _
             "AND (DBBatidos.CodGrupo <> 7 OR DBBatidos.CodGrupo Is Null) " & _
             "AND SubReceta.codporcion Is Null " & _
             "AND Grupos.control = True " & _
             "GROUP BY Left(DBIngredientes.CodIngrediente,1), DBIngredientes.Tipo, " & _
             "DBIngredientes.Nombre, SubReceta.codporcion, SubReceta.CodIngrediente, " & _
             "DBIngredientes.Unidad " & _
             "ORDER BY Left(DBIngredientes.CodIngrediente,1), DBIngredientes.Tipo, DBIngredientes.Nombre;"
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No se encontraron registros para exportar.", vbInformation
        GoTo Limpiar
    End If
    
    ' Posición inicial para datos
    fila = 2
    
    ' Obtener total de registros
    Dim totalRegistros As Long
    rs.MoveLast
    totalRegistros = rs.RecordCount
    rs.MoveFirst
    
    ' Recorrer el recordset y llenar datos
    Do While Not rs.EOF
        ' Actualizar mensaje de progreso
        If (fila - 1) Mod 25 = 0 Then
            DoCmd.Echo False, "Procesando registro " & (fila - 1) & " de " & totalRegistros & "..."
            DoEvents
        End If
        
        ' Construir nombre del producto: [Nombre] & Iif([Unidad]="gr","oz",[Unidad])
        On Error Resume Next
        If rs!Unidad = "gr" Then
            NombreProducto = rs!Nombre & " oz"
        Else
            NombreProducto = rs!Nombre & " " & rs!Unidad
        End If
        
        If Err.Number <> 0 Then
            xlSheet.Cells(fila, 1).Value = "Error en " & rs!CodIngrediente
            Err.Clear
        Else
            xlSheet.Cells(fila, 1).Value = NombreProducto
        End If
        On Error GoTo 0
        
        ' Determinar divisor: SiInm([Unidad]="gr",28.375,1)
        On Error Resume Next
        If rs!Unidad = "gr" Then
            divisor = 28.375
        Else
            divisor = 1
        End If
        On Error GoTo 0
        
        ' Calcular consumos para cada semana
        For j = 1 To numSemanas
            semana_actual = semana_desde + j - 1
            
            On Error Resume Next
            ' Calcular: VentasXTipoNoPorcionSemana(semana, CodIngrediente) / divisor
            resultado = VentasXTipoNoPorcionSemana(semana_actual, rs!CodIngrediente)
            
            If Err.Number <> 0 Then
                xlSheet.Cells(fila, j + 1).Value = 0
                Err.Clear
            Else
                If IsNumeric(resultado) And divisor <> 0 Then
                    xlSheet.Cells(fila, j + 1).Value = resultado / divisor
                Else
                    xlSheet.Cells(fila, j + 1).Value = 0
                End If
            End If
            On Error GoTo 0
        Next j
        
        fila = fila + 1
        rs.MoveNext
    Loop
    
    ' Aplicar formato
    If fila > 2 Then
        With xlSheet
            ' Autoajustar columna de productos
            .Columns("A").EntireColumn.AutoFit
            If .Columns("A").ColumnWidth > 80 Then .Columns("A").ColumnWidth = 80
            
            ' Formato de números para las columnas de consumos
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).NumberFormat = "#,##0.00"
            
            ' Agregar bordes
            .Range(.Cells(1, 1), .Cells(fila - 1, numSemanas + 1)).Borders.LineStyle = 1
            
            ' Filas con color alternado
            For I = 2 To fila - 1 Step 2
                .Range(.Cells(I, 1), .Cells(I, numSemanas + 1)).Interior.color = RGB(240, 240, 240)
            Next I
            
            ' Alinear datos a la derecha
            .Range(.Cells(2, 2), .Cells(fila - 1, numSemanas + 1)).horizontalAlignment = -4152
        End With
    End If
    
    ' --- Guardar archivo ---
    
    ' Construir la ruta del directorio
    On Error Resume Next
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
    
    ' Verificar si el directorio existe, si no, crearlo
    If Dir(Direccion, vbDirectory) = "" Then
        MkDir Direccion
    End If
    
    ' Construir el nombre del archivo: [nombrelocal()] - Sem [semana_desde] hasta [semana_hasta] No Porciones Compacto.xlsx
    archivo = Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - Sem " & semana_desde & " hasta " & semana_hasta & " - No Porciones Compacto.xlsx"
    
    ' Guardar el libro
    xlBook.SaveAs archivo, 51
    
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo: " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Cerrar Excel sin guardar cambios adicionales
    xlBook.Close False
    xlApp.Quit
    
    ' Restaurar cursor y mensajes
    DoCmd.Hourglass False
    DoCmd.Echo True
    
Limpiar:
    ' Limpiar objetos
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True
    On Error GoTo 0
    Exit Sub

ErrorExcel:
    MsgBox "Error al crear Excel. Asegúrate de tener Microsoft Excel instalado.", vbCritical
    Resume Limpiar
End Sub



Private Sub Comando340_Click()

    ' Declaraciones
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object ' Excel.Application
    Dim xlBook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim strSQL As String
    Dim I As Long, j As Long
    Dim semana_desde As Integer
    Dim semana_hasta As Integer
    Dim numSemanas As Integer
    Dim fila As Long
    Dim semana_actual As Integer
    Dim resultado As Variant
    Dim cantidadpaquete As Variant
    Dim startTime As Single
    Dim dict As Object ' Para verificar duplicados (opcional)
    Dim clave As String
    
    ' Declaraciones para el manejo de archivos
    Dim Direccion As String
    Dim archivo As String
    
    startTime = Timer ' Para medir tiempo de ejecución
    
    ' Solicitar rango de semanas
    semana_desde = Me.semana_desde
    semana_hasta = Me.semana_hasta
    
    ' Validar entrada
    If semana_desde <= 0 Or semana_hasta < semana_desde Then
        MsgBox "Rango de semanas inválido. La semana inicial debe ser mayor a 0 y la semana final debe ser mayor o igual a la inicial.", vbExclamation
        Exit Sub
    End If
    
    numSemanas = semana_hasta - semana_desde + 1
    
    ' Mostrar mensaje de progreso
    DoCmd.Hourglass True
    DoCmd.Echo False, "Generando reporte de consumos MOSTRADOR. Por favor espere..."
    
    ' Crear instancia de Excel INVISIBLE
    On Error GoTo ErrorExcel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error GoTo 0
    
    ' Configurar la hoja
    With xlSheet
        .Name = "Mostrador Sem " & semana_desde & "-" & semana_hasta
        
        ' Encabezados
        .Cells(1, 1).Value = "Marca"
        .Cells(1, 2).Value = "Producto"
        .Cells(1, 3).Value = "Cantidad Paquete"
        
        ' Encabezados de semanas (a partir de columna 4)
        For I = 1 To numSemanas
            .Cells(1, I + 3).Value = "Semana " & (semana_desde + I - 1)
            .Cells(1, I + 3).Font.Bold = True
            .Cells(1, I + 3).Interior.color = RGB(200, 200, 200)
            .Cells(1, I + 3).horizontalAlignment = -4108 ' xlCenter
            .Cells(1, I + 3).ColumnWidth = 15
        Next I
        
        ' Formato de encabezados fijos
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.color = RGB(200, 200, 200)
        .Columns("A").ColumnWidth = 15  ' Marca
        .Columns("B").ColumnWidth = 40  ' Producto
        .Columns("C").ColumnWidth = 15  ' Cantidad Paquete
        
        ' Título del reporte
        .Rows("1:1").RowHeight = 20
    End With
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    ' Consulta SQL COMPLETA con GROUP BY (tal como estaba originalmente)
    
    strSQL = "SELECT " & _
             "DBBatidos.Marca, " & _
             "DLookUp('[PaquetePorciones]','[Cotizaciones]','[CodCotizacion]=' & CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido])) AS cantidadpaquete, " & _
             "DBBatidos.CodGrupo, " & _
             "DBBatidos.Vigencia, " & _
             "SubReceta.InsumoClave, " & _
             "IIf(IsNull([SubReceta]![codporcion]), " & _
             "DLookUp('[Nombre]','[DBIngredientes]','[CodIngrediente]=''' & [SubReceta]![CodIngrediente] & ''''), " & _
             "nombreproductocotiprocesado([SubReceta]![codporcion])) AS nombreprod, " & _
             "CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido]) AS coti, " & _
             "SubReceta.codporcion, " & _
             "SubReceta.CodIngrediente " & _
             "FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido " & _
             "GROUP BY DBBatidos.Marca, " & _
             "DLookUp('[PaquetePorciones]','[Cotizaciones]','[CodCotizacion]=' & CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido])), " & _
             "DBBatidos.CodGrupo, " & _
             "DBBatidos.Vigencia, " & _
             "SubReceta.InsumoClave, " & _
             "IIf(IsNull([SubReceta]![codporcion]), " & _
             "DLookUp('[Nombre]','[DBIngredientes]','[CodIngrediente]=''' & [SubReceta]![CodIngrediente] & ''''), " & _
             "nombreproductocotiprocesado([SubReceta]![codporcion])), " & _
             "CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido]), SubReceta.codporcion, " & _
             "SubReceta.CodIngrediente HAVING DBBatidos.CodGrupo = 7 AND DBBatidos.Vigencia = True " & _
             "AND SubReceta.InsumoClave = True ORDER BY DBBatidos.Marca, " & _
             "DLookUp('[PaquetePorciones]','[Cotizaciones]','[CodCotizacion]=' & CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido]));"
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No se encontraron registros para exportar.", vbInformation
        GoTo Limpiar
    End If
    
    ' Posición inicial para datos
    fila = 2
    
    ' Obtener total de registros
    Dim totalRegistros As Long
    rs.MoveLast
    totalRegistros = rs.RecordCount
    rs.MoveFirst
    
    ' Opcional: Crear diccionario para verificar duplicados (solo para debug)
    ' Set dict = CreateObject("Scripting.Dictionary")
    
    ' Recorrer el recordset y llenar datos
    Do While Not rs.EOF
        ' Actualizar mensaje de progreso
        If (fila - 1) Mod 25 = 0 Then
            DoCmd.Echo False, "Procesando registro " & (fila - 1) & " de " & totalRegistros & "..."
            DoEvents
        End If
        
        ' Crear clave única para verificar duplicados (opcional)
        ' clave = rs!Marca & "|" & rs!nombreprod & "|" & rs!codporcion & "|" & rs!CodIngrediente
        ' If dict.exists(clave) Then
        '     rs.MoveNext
        '     GoTo Siguiente
        ' Else
        '     dict.Add clave, True
        ' End If
        
        ' Marca (columna A)
        On Error Resume Next
        xlSheet.Cells(fila, 1).Value = rs!Marca
        On Error GoTo 0
        
        ' Nombre del producto (columna B) - usando nombreprod
        On Error Resume Next
        xlSheet.Cells(fila, 2).Value = rs!nombreprod
        On Error GoTo 0
        
        ' Cantidad paquete (columna C)
        On Error Resume Next
        cantidadpaquete = rs!cantidadpaquete
        If IsNull(cantidadpaquete) Or Not IsNumeric(cantidadpaquete) Then
            xlSheet.Cells(fila, 3).Value = 0
        Else
            xlSheet.Cells(fila, 3).Value = cantidadpaquete
        End If
        On Error GoTo 0
        
        ' Calcular consumos para cada semana (a partir de columna D)
        For j = 1 To numSemanas
            semana_actual = semana_desde + j - 1
            
            On Error Resume Next
            ' Determinar qué función usar según si tiene codporcion o no
            If IsNull(rs!codporcion) Then
                ' Usar VentasXTipoNoPorcionSemana
                resultado = VentasXTipoNoPorcionSemana(semana_actual, rs!CodIngrediente)
            Else
                ' Usar VentasXTipoPorcionSemana
                resultado = VentasXTipoPorcionSemana(semana_actual, rs!codporcion)
            End If
            
            If Err.Number <> 0 Then
                xlSheet.Cells(fila, j + 3).Value = 0
                Err.Clear
            Else
                If IsNumeric(resultado) Then
                    xlSheet.Cells(fila, j + 3).Value = resultado
                Else
                    xlSheet.Cells(fila, j + 3).Value = 0
                End If
            End If
            On Error GoTo 0
        Next j
        
        fila = fila + 1
' Siguiente:
        rs.MoveNext
    Loop
    
    ' Aplicar formato
    If fila > 2 Then
        With xlSheet
            ' Autoajustar columnas
            .Columns("A").EntireColumn.AutoFit
            .Columns("B").EntireColumn.AutoFit
            .Columns("C").EntireColumn.AutoFit
            
            ' Limitar ancho de columna de productos si es muy grande
            If .Columns("B").ColumnWidth > 80 Then .Columns("B").ColumnWidth = 80
            
            ' Formato de números para las columnas de consumos (a partir de columna D)
            .Range(.Cells(2, 4), .Cells(fila - 1, numSemanas + 3)).NumberFormat = "#,##0.00"
            
            ' Formato para cantidad paquete (columna C)
            .Range(.Cells(2, 3), .Cells(fila - 1, 3)).NumberFormat = "#,##0"
            
            ' Agregar bordes a toda la tabla
            .Range(.Cells(1, 1), .Cells(fila - 1, numSemanas + 3)).Borders.LineStyle = 1
            
            ' Filas con color alternado
            For I = 2 To fila - 1 Step 2
                .Range(.Cells(I, 1), .Cells(I, numSemanas + 3)).Interior.color = RGB(240, 240, 240)
            Next I
            
            ' Alinear datos numéricos a la derecha
            .Range(.Cells(2, 3), .Cells(fila - 1, numSemanas + 3)).horizontalAlignment = -4152
        End With
    End If
    
    ' --- Guardar archivo ---
    
    ' Construir la ruta del directorio
    On Error Resume Next
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
    
    ' Verificar si el directorio existe, si no, crearlo
    If Dir(Direccion, vbDirectory) = "" Then
        MkDir Direccion
    End If
    
    ' Construir el nombre del archivo
    archivo = Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - Sem " & semana_desde & " hasta " & semana_hasta & " - Mostrador.xlsx"
    
    ' Guardar el libro
    xlBook.SaveAs archivo, 51
    
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo: " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Cerrar Excel sin guardar cambios adicionales
    xlBook.Close False
    xlApp.Quit
    
    ' Restaurar cursor y mensajes
    DoCmd.Hourglass False
    DoCmd.Echo True

Limpiar:
    ' Limpiar objetos
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    ' If Not dict Is Nothing Then Set dict = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True
    On Error GoTo 0
    Exit Sub

ErrorExcel:
    MsgBox "Error al crear Excel. Asegúrate de tener Microsoft Excel instalado.", vbCritical
    Resume Limpiar
End Sub
    
Private Sub Comando350_Click()
    ' Declaraciones
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object ' Excel.Application
    Dim xlBook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim strSQL As String
    Dim I As Long, j As Long
    Dim semana_desde As Integer
    Dim semana_hasta As Integer
    Dim numSemanas As Integer
    Dim fila As Long
    Dim semana_actual As Integer
    Dim resultado As Variant
    Dim NombreProducto As String
    Dim startTime As Single
    
    ' Declaraciones para el manejo de archivos
    Dim Direccion As String
    Dim archivo As String
    
    startTime = Timer ' Para medir tiempo de ejecución
    
    ' Solicitar rango de semanas
    semana_desde = Me.semana_desde
    semana_hasta = Me.semana_hasta
    
    ' Validar entrada
    If semana_desde <= 0 Or semana_hasta < semana_desde Then
        MsgBox "Rango de semanas inválido. La semana inicial debe ser mayor a 0 y la semana final debe ser mayor o igual a la inicial.", vbExclamation
        Exit Sub
    End If
    
    numSemanas = semana_hasta - semana_desde + 1
    
    ' Mostrar mensaje de progreso
    DoCmd.Hourglass True
    DoCmd.Echo False, "Generando reporte de PORCIONES COMPACTO. Por favor espere..."
    
    ' Crear instancia de Excel INVISIBLE
    On Error GoTo ErrorExcel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error GoTo 0
    
    ' Configurar la hoja
    With xlSheet
        .Name = "Porciones Sem " & semana_desde & "-" & semana_hasta
        
        ' Encabezados
        .Cells(1, 1).Value = "Almacén"
        .Cells(1, 2).Value = "Tipo"
        .Cells(1, 3).Value = "Producto"
        
        ' Encabezados de semanas (a partir de columna 4)
        For I = 1 To numSemanas
            .Cells(1, I + 3).Value = "Semana " & (semana_desde + I - 1)
            .Cells(1, I + 3).Font.Bold = True
            .Cells(1, I + 3).Interior.color = RGB(200, 200, 200)
            .Cells(1, I + 3).horizontalAlignment = -4108 ' xlCenter
            .Cells(1, I + 3).ColumnWidth = 15
        Next I
        
        ' Formato de encabezados fijos
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.color = RGB(200, 200, 200)
        .Columns("A").ColumnWidth = 15  ' Almacén
        .Columns("B").ColumnWidth = 15  ' Tipo
        .Columns("C").ColumnWidth = 40  ' Producto
        
        ' Título del reporte
        .Rows("1:1").RowHeight = 20
    End With
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    
    ' Consulta SQL completa con GROUP BY para Porciones Compacto
    strSQL = "SELECT " & _
             "DLookUp('[CodAlmacenamiento]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) AS almacen, " & _
             "DBIngredientes.Tipo, " & _
             "DBIngredientes.Nombre, " & _
             "SubReceta.codporcion, " & _
             "PorcionDentroDeMezcla([SubReceta]![codporcion]) AS mezcla, " & _
             "Grupos.control, " & _
             "DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) AS Descontinuado " & _
             "FROM ((SubReceta " & _
             "INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) " & _
             "INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente) " & _
             "INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo " & _
             "GROUP BY " & _
             "DLookUp('[CodAlmacenamiento]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]), " & _
             "DBIngredientes.Tipo, DBIngredientes.Nombre, [DBBatidos]![CodGrupo]<>7, " & _
             "SubReceta.codporcion, PorcionDentroDeMezcla([SubReceta]![codporcion]), Grupos.control, " & _
             "DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) " & _
             "HAVING ([DBBatidos]![CodGrupo]<>7) = True AND SubReceta.codporcion Is Not Null " & _
             "AND PorcionDentroDeMezcla([SubReceta]![codporcion]) = 0 AND Grupos.control = True " & _
             "AND DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) = False ORDER BY " & _
             "DLookUp('[CodAlmacenamiento]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]), " & _
             "DBIngredientes.Tipo, DBIngredientes.Nombre;"
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No se encontraron registros para exportar.", vbInformation
        GoTo Limpiar
    End If
    
    ' Posición inicial para datos
    fila = 2
    
    ' Obtener total de registros
    Dim totalRegistros As Long
    rs.MoveLast
    totalRegistros = rs.RecordCount
    rs.MoveFirst
    
    ' Recorrer el recordset y llenar datos
    Do While Not rs.EOF
        ' Actualizar mensaje de progreso
        If (fila - 1) Mod 25 = 0 Then
            DoCmd.Echo False, "Procesando registro " & (fila - 1) & " de " & totalRegistros & "..."
            DoEvents
        End If
        
        ' Almacén (columna A)
        On Error Resume Next
        xlSheet.Cells(fila, 1).Value = rs!Almacen
        On Error GoTo 0
        
        ' Tipo (columna B)
        On Error Resume Next
        xlSheet.Cells(fila, 2).Value = rs!Tipo
        On Error GoTo 0
        
        ' Nombre del producto (columna C) - combinando las dos funciones
        On Error Resume Next
        NombreProducto = nombreproductocotiprocesadosinunidad(rs!codporcion) & _
                         nombreproductocotiprocesadosolounidad(rs!codporcion)
        
        If Err.Number <> 0 Then
            xlSheet.Cells(fila, 3).Value = "Error en " & rs!codporcion
            Err.Clear
        Else
            xlSheet.Cells(fila, 3).Value = NombreProducto
        End If
        On Error GoTo 0
        
        ' Calcular consumos para cada semana (a partir de columna D)
        For j = 1 To numSemanas
            semana_actual = semana_desde + j - 1
            
            On Error Resume Next
            ' Usar consumoporciones(codporcion, semana)
            resultado = consumoporciones(rs!codporcion, semana_actual)
            
            If Err.Number <> 0 Then
                xlSheet.Cells(fila, j + 3).Value = 0
                Err.Clear
            Else
                If IsNumeric(resultado) Then
                    xlSheet.Cells(fila, j + 3).Value = resultado
                Else
                    xlSheet.Cells(fila, j + 3).Value = 0
                End If
            End If
            On Error GoTo 0
        Next j
        
        fila = fila + 1
        rs.MoveNext
    Loop
    
    ' Aplicar formato
    If fila > 2 Then
        With xlSheet
            ' Autoajustar columnas
            .Columns("A").EntireColumn.AutoFit
            .Columns("B").EntireColumn.AutoFit
            .Columns("C").EntireColumn.AutoFit
            
            ' Limitar ancho de columna de productos si es muy grande
            If .Columns("C").ColumnWidth > 80 Then .Columns("C").ColumnWidth = 80
            
            ' Formato de números para las columnas de consumos (a partir de columna D)
            .Range(.Cells(2, 4), .Cells(fila - 1, numSemanas + 3)).NumberFormat = "#,##0.00"
            
            ' Agregar bordes a toda la tabla
            .Range(.Cells(1, 1), .Cells(fila - 1, numSemanas + 3)).Borders.LineStyle = 1
            
            ' Filas con color alternado
            For I = 2 To fila - 1 Step 2
                .Range(.Cells(I, 1), .Cells(I, numSemanas + 3)).Interior.color = RGB(240, 240, 240)
            Next I
            
            ' Alinear datos numéricos a la derecha
            .Range(.Cells(2, 4), .Cells(fila - 1, numSemanas + 3)).horizontalAlignment = -4152
            
            ' Centrar texto de Almacén y Tipo
            .Range(.Cells(2, 1), .Cells(fila - 1, 2)).horizontalAlignment = -4108 ' xlCenter
        End With
    End If
    
    ' --- Guardar archivo ---
    
    ' Construir la ruta del directorio
    On Error Resume Next
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
    
    ' Verificar si el directorio existe, si no, crearlo
    If Dir(Direccion, vbDirectory) = "" Then
        MkDir Direccion
    End If
    
    ' Construir el nombre del archivo: [nombrelocal()] - Sem [semana_desde] hasta [semana_hasta] Porciones Compacto.xlsx
    archivo = Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - Sem " & semana_desde & " hasta " & semana_hasta & " - Porciones Compacto.xlsx"
    
    ' Guardar el libro
    xlBook.SaveAs archivo, 51
    
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo: " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Cerrar Excel sin guardar cambios adicionales
    xlBook.Close False
    xlApp.Quit
    
    ' Restaurar cursor y mensajes
    DoCmd.Hourglass False
    DoCmd.Echo True

Limpiar:
    ' Limpiar objetos
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True
    On Error GoTo 0
    Exit Sub

ErrorExcel:
    MsgBox "Error al crear Excel. Asegúrate de tener Microsoft Excel instalado.", vbCritical
    Resume Limpiar
End Sub

Private Sub Comando352_Click()
Call Comando350_Click
Call Comando334_Click
Call Comando340_Click
MsgBox "Descarga de Escel completo de sucursal actual"
End Sub

Private Sub Comando357_Click()

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Main_DB2.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim dbexterna As DAO.Database
Dim rst As DAO.Recordset
Dim miSQL As String
Set dbexterna = DBEngine.OpenDatabase(h)
miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
Set rst = dbexterna.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF

    If CurrentProject.Name = "Pitaya_System.accdb" Then 'archivo base de sistema
            Call actualizartablas(1, rst("CodLocal"))
    Else
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        Exit Sub
    End If

    Call Comando350_Click
    Call Comando334_Click
    Call Comando340_Click
    
    rst.MoveNext
Loop
dbexterna.Close
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "Excel Generado de Todas las sucurusales"
Exit Sub

Nulo:
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "No se han generado los Excel correctamente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
End Sub
