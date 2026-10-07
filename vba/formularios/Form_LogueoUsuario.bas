' ==========================================================
' Modulo  : Form_LogueoUsuario
' Tipo    : 100
' Lineas  : 51
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database




Private Sub Comando48_Click()
Dim userx As Integer
'Call ObtenerMarcacionesHoy(codigolocal())
Call DescargarTablaCompleta("marcaciones", "marcacioneshoy", "sucursal_codigo=" & codigoLocal() & " AND fecha='" & Format(Now, "yyyy-mm-dd") & "'")

If IsNull(Me.password) Or Me.password = 0 Then
    'no hace nada
    MsgBox "Ingresar contraseña"
Else
    userx = ObtenerCodOperario(Me.password) 'buscar usuarios de la tabla de marcaciones hoy , si no marcacron no sale
    If userx = 0 Then 'probar con los que tienne permiso de lider
        userx = usuarioconpermisosclave(Me.password)
    End If
    
    If userx <> 0 Then
        Select Case Me.origenlogueo
            Case "[Introduccion]"
                'Call EnviarMensajeTelegram("Abriendo Facturacion '" & NombreOperario(userx) & "'", grupotgerencia())
                DoCmd.OpenForm "Main Pitaya"
                [Forms]![Main Pitaya]![codigologin] = userx
                [Forms]![Main Pitaya].Form.Requery
                DoCmd.Close acForm, "Introduccion"

            Case Else ' resto de formularios
                DoCmd.OpenForm Me.origenlogueo
                Forms(Me.origenlogueo).Controls("codigologin") = userx
                
        End Select
        
        DoCmd.Close acForm, "LogueoUsuario"
        
    Else
        MsgBox "Clave Erronea", vbOKOnly, "CLAVE ERRONEA"
        Me.password = ""
        Exit Sub
    End If
End If


End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

