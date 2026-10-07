' ==========================================================
' Modulo  : Form_Check In/Out
' Tipo    : 100
' Lineas  : 137
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub CodigoBusqueda_Exit(Cancel As Integer)
On Error GoTo Nulo
If IsNull(horaingresopendienteoperariodia(Me.CodigoBusqueda, Date)) Or horaingresopendienteoperariodia(Me.CodigoBusqueda, Date) = 0 Then
    Me.Ingreso = ""
    Me.Comando48.Caption = "MARCAR ENTRADA"
    Me.tipomarcacion = 1 'marcar entrada nuevo registro
    Me.registroencurso = 0
Else
    Me.Ingreso = horaingresopendienteoperariodia(Me.CodigoBusqueda, Date)
    Me.Comando48.Caption = "MARCAR SALIDA"
    Me.tipomarcacion = 2 ' marcar salida ya existe registro
    Me.registroencurso = registroencursooperariodia(Me.CodigoBusqueda, Date)
End If
Exit Sub

Nulo:
Me.Comando48.Caption = "MARCAR ENTRADA/SALIDA"
End Sub

Private Sub Comando149_Click()
If Me.password = CorroborarClave(Me.CodigoBusqueda) Then
    DoCmd.OpenForm "Historial Asistencia", , , "[CodOperario]=" & Me.CodigoBusqueda
Else
    MsgBox "Clave Erronea", vbOKOnly, "CLAVE ERRONEA"
End If
End Sub



Private Sub Comando48_Click()

If Me.password = CorroborarClave(Me.CodigoBusqueda) Then
    If Me.tipomarcacion = 2 Then ' MARCAR SALIDA
        If (Time - Me.Ingreso) * 24 * 60 < 120 Then
            MsgBox "No puede marcar doble entrada"
            Exit Sub
        End If
        '''' Ya no aplica solo busca fechas marcacion del dia
        'If Me.Fecha = Date Then ' tiene marcacion de ingres dentro de la fecha que va a marcar
        '    Me.Salida = Time
        '    MsgBox "Acaba de marcar salida " & Me.Salida
        'Else ' hay amrcacion de ingreso de fecha anterior
        '    Me.Salida = Time
        '    MsgBox "SE HA MARCADO SALIDA DE FECHA ANTERIOR, MARCAR NUEVAMENTE INGRESO", vbOKOnly, "FALTANTE DE FECHA DE SALIDA"
        'End If
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE RegistroHorario SET Salida = #" & Time & "# WHERE CodHorario = " & Me.registroencurso
        DoCmd.SetWarnings True
        
        MsgBox "Acaba de Marcar Salida " & Time, vbOKOnly, "Salida"
        
        'mensaje de salida
        'Call EnviarMensajeTelegram("Salida de " & NombreOperario(Me.CodigoBusqueda) & " " & CStr(Time) & "", grupotgerencia())
        'Foto de salida
        'Call FotoTemporalTelegram
        
        
    Else ' MARCAR ENTRADA
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO RegistroHorario(Ingreso, Fecha, CodOperario) values" & _
        " (#" & Time & "#, #" & Date & "#, " & Me.CodigoBusqueda & ")"
        DoCmd.SetWarnings True
        
        MsgBox "Acaba de Marcar Ingreso " & Time, vbOKOnly, "Ingreso"
        
        'Mensaje de entrada
        'Call EnviarMensajeTelegram("Ingreso de " & NombreOperario(Me.CodigoBusqueda) & " " & CStr(Time) & "", grupotgerencia())
        'Foto de entrada
        'Call FotoTemporalTelegram
    End If
    
    Me.CodigoBusqueda = ""
    Me.password = ""
    Me.Comando48.Caption = "MARCAR ENTRADA/SALIDA"
    Me.Ingreso = ""
    Call ListaOperariosdeTurno
    
    If codigoLocal() <> 9 And codigoLocal() <> 13 And codigoLocal() <> 0 Then
        On Error Resume Next
        Call FotoOperarioCheck(Me.CodigoBusqueda, Me.tipomarcacion)
    End If
Else
    MsgBox "Clave Erronea", vbOKOnly, "CLAVE ERRONEA"
    Me.password = ""
    Exit Sub
End If

'Dim Direccion As String: Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal()
'If Dir(Direccion, vbDirectory) = "" Then MkDir Direccion

'Me.CodigoBusqueda = Nulo
'Me.Requery
'Call ListaOperariosdeTurno

'Dim ruta As String: ruta = Direccion & "\De Turno.txt"
'Open ruta For Output As #1
'Print #1, Me.deturno
'Close #1
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

If CurrentProject.Name <> "ModuloRRHH.accdb" Then
    Call importartablasweb
End If
Call ListaOperariosdeTurno
End Sub

Private Sub ListaOperariosdeTurno()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim I As Integer

Me.deturno = ""
I = 1
miSQL = "SELECT Operarios.Nombre, Operarios.Apellido, RegistroHorario.Ingreso, RegistroHorario.Salida, RegistroHorario.Fecha" & _
" FROM RegistroHorario INNER JOIN Operarios ON RegistroHorario.CodOperario = Operarios.CodOperario" & _
" WHERE (((RegistroHorario.Salida) Is Null) AND ((RegistroHorario.Fecha)=Date()))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    Me.deturno = Me.deturno & IIf(I = 0, Chr(13) & Chr(10), "") & rst("Nombre") & "  " & rst("Apellido") & " --> " & rst("Ingreso")
    rst.MoveNext
    I = 0
Loop
rst.Close
Exit Sub

Nulo:
End Sub

