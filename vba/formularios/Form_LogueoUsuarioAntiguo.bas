' ==========================================================
' Modulo  : Form_LogueoUsuarioAntiguo
' Tipo    : 100
' Lineas  : 50
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:30
' ==========================================================
Option Compare Database


Private Sub Comando48_Click()
Dim userx As Integer

If IsNull(Me.CodigoBusqueda) Or Me.CodigoBusqueda = 0 Then
    'no hace nada
    MsgBox "Colaborador no puede estar vacio"
Else
    
    If IsNull(Me.password) Or Me.password = 0 Then
        'no hace nada
        MsgBox "Ingresar contraseña"
    Else
        userx = Me.CodigoBusqueda

        If Me.password = CorroborarClave(Me.CodigoBusqueda) Or Me.password = CorroborarClave(cargooperariooperativo(15, 0, Date)) Then
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
            
            DoCmd.Close acForm, "LogueoUsuarioAntiguo"
            
        Else
            MsgBox "Clave Erronea", vbOKOnly, "CLAVE ERRONEA"
            Me.password = ""
            Exit Sub
        End If
    End If
End If


End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

