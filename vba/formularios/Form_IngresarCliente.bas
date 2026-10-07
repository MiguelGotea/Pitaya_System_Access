' ==========================================================
' Modulo  : Form_IngresarCliente
' Tipo    : 100
' Lineas  : 135
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database


Private Sub ActualizarFechaCumple()
    Dim Dia As Integer
    Dim mesTexto As String
    Dim mesNum As Integer
    Dim anio As Integer

    ' Solo procede si las 3 celdas tienen datos (ya validados individualmente)
    If Not IsNull(Me.adia) And Me.adia <> "" And _
       Not IsNull(Me.ames) And Me.ames <> "" And _
       Not IsNull(Me.aano) And Me.aano <> "" Then

        Dia = Me.adia
        mesTexto = Me.ames
        anio = Me.aano

        mesNum = ObtenerNumeroMes(mesTexto)

        If mesNum > 0 Then
            On Error Resume Next
            Me.acumple = DateSerial(anio, mesNum, Dia)
            If Err.Number <> 0 Then
                MsgBox "La fecha ingresada no es válida (por ejemplo, 31 de Febrero).", vbExclamation
                Me.acumple = Null
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Else
        ' Si alguna celda quedó vacía, no se toca acumple (o puedes limpiarlo si prefieres)
        Me.acumple = Null
    End If
End Sub

Private Function ObtenerNumeroMes(ByVal Mes As String) As Integer
    Select Case Mes
        Case "Enero":       ObtenerNumeroMes = 1
        Case "Febrero":     ObtenerNumeroMes = 2
        Case "Marzo":       ObtenerNumeroMes = 3
        Case "Abril":       ObtenerNumeroMes = 4
        Case "Mayo":        ObtenerNumeroMes = 5
        Case "Junio":       ObtenerNumeroMes = 6
        Case "Julio":       ObtenerNumeroMes = 7
        Case "Agosto":      ObtenerNumeroMes = 8
        Case "Setiembre":   ObtenerNumeroMes = 9
        Case "Octubre":     ObtenerNumeroMes = 10
        Case "Noviembre":   ObtenerNumeroMes = 11
        Case "Diciembre":   ObtenerNumeroMes = 12
        Case Else:          ObtenerNumeroMes = 0
    End Select
End Function
Private Sub aano_Exit(Cancel As Integer)

If Not IsNull(Me.aano) Then

    If CInt(Me.aano) < 1900 Or CInt(Me.aano) > Year(Date) - 10 Then
        MsgBox "Ingresar un numero correcto de año"
        Me.aano = ""
    End If
End If
Call ActualizarFechaCumple
End Sub

Private Sub adia_Exit(Cancel As Integer)
If Not IsNull(Me.adia) Then
    If CInt(Me.adia) > 31 Then
        MsgBox "Ingresar un numero correcto de dia"
        Me.adia = ""
    End If
End If
Call ActualizarFechaCumple
End Sub

Private Sub ames_Exit(Cancel As Integer)
Call ActualizarFechaCumple
End Sub

Private Sub Comando36_Click()
If Me.acodigo = "" Or IsNull(Me.acodigo) Then
    MsgBox "Ingresar numero de la membresia"
    Exit Sub
End If
If Me.anombre = "" Or IsNull(Me.anombre) Then
    MsgBox "Ingresar un nombre "
    Exit Sub
End If

If Me.aapellido = "" Or IsNull(Me.aapellido) Then
    MsgBox "Ingresar un apellido"
    Exit Sub
End If
If Me.acelular = "" Or IsNull(Me.acelular) Then
    MsgBox "Ingresar un numero de celular"
    Exit Sub
End If
If Me.acumple = "" Or IsNull(Me.acumple) Then
    MsgBox "Ingresar fecha de nacimiento"
    Exit Sub
End If


DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO ClientesClub(CodCliente, Nombre, Apellidos, Celular, Cumpleanos, Correo, [Fecha de Inscripcion], local)" & _
" values (" & Me.acodigo & ", '" & Me.anombre & "', '" & Me.aapellido & "', '" & Me.acelular & "', #" & Me.acumple & "#," & _
" '" & Me.acorreo & "', #" & Me.afecha & "#, '" & Me.alocal & "')"
DoCmd.SetWarnings True

      
MsgBox "Membresia registrada correctamente"

If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    [Forms]![Nota de Pedido]![CodCliente] = Me.acodigo
    [Forms]![Nota de Pedido]![Nombre Provicional] = Me.anombre
    [Forms]![Nota de Pedido]![Nombre Provicional].Enabled = False
    
    [Forms]![Nota de Pedido]![Encabezado_automático0] = "NOTA DE PEDIDO " & Me.acodigo & " - " & UCase(Me.anombre)
    [Forms]![Nota de Pedido]![Cuadro_combinado123] = UCase(Me.anombre)
    [Forms]![Nota de Pedido]![ppedido].Requery
    [Forms]![Nota de Pedido]![piniciales] = 0
    [Forms]![Nota de Pedido]![pacumulado] = 0
    [Forms]![Nota de Pedido]![clubguardado] = Me.acodigo
End If
DoCmd.Close acForm, "IngresarCliente"
End Sub

Private Sub Comando37_Click()
DoCmd.Close acForm, "IngresarCliente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub

