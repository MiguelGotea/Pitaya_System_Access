' ==========================================================
' Modulo  : Form_Introduccion
' Tipo    : 100
' Lineas  : 261
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Private Sub starplayer_Click()
'If Not ProgramaEstaAbierto("vlc.exe") Then
'    Dim archivomusica As String
'    Dim rutavlc As String
'
'    archivomusica = "C:\Users\" & NombreSistema() & "\Google Drive BP\Pitaya Music\youtube\youtube" & codigolocal() & ".m3u8"
'    rutavlc = "C:\Program Files\VideoLAN\VLC\vlc.exe"
'    Call Shell(rutavlc & " " & """" & archivomusica & """", vbMinimizedNoFocus)
'End If
'Exit Sub ' usaremos archivo vlc
Dim ruta As String
ruta = "C:\Users\" & NombreSistema() & "\Google Drive BP\Pitaya Music\"
Set fso = CreateObject("Scripting.FileSystemObject")
Set carpeta = fso.GetFolder(ruta)

Randomize
I = Int(Rnd() * carpeta.Files.Count + 1)
j = 1

For Each archivo In carpeta.Files
    If j = I Then
        If archivo.Type = "MP3 Format Sound" Or archivo.Type = "Sonido en formato MP3" Then
            automusica.url = ruta & archivo.Name
            cancion.Value = Left(archivo.Name, 20)
            automusica.Object.Controls.play
            automusica.settings.volume = DLookup("[VolumenMusica]", "DatosSistema")
        End If
    End If
    j = j + 1
Next archivo

Set fso = Nothing
Set carpeta = Nothing
Set ficheros = Nothing
End Sub

Private Sub Comando1217_Click()
'cerrar todo
If Weekday(Date, 2) = 7 And Hour(Now) > 19 Then
    MsgBox "Los domingos no se puede cerrar el sistema, ni apagar la computadora"
    Exit Sub
End If

If Weekday(Date, 2) = 7 And Hour(Now) > 12 Then
    If MsgBox("Los domingos no se tiene que apagar la computadora y dejar el sistema abierto, desea cerrar de todas maneras?", vbYesNo, "Alerta de Domingo") = vbNo Then
        Exit Sub
    Else
    End If
End If


[Forms]![Introduccion].Comando1437.Enabled = False 'Gestion
[Forms]![Introduccion].Comando1438.Enabled = False 'Ventas
[Forms]![Introduccion].Comando1217.Enabled = False 'Gestion
[Forms]![Introduccion].senuelo.Visible = True
'Call CreateObject("WScript.Shell").Run("taskkill /f /im " & "vlc.exe", 0, True) 'Cerrar V


DoCmd.Quit
End Sub


Private Sub Comando1437_Click()
DoCmd.OpenForm "IngresoClaveAdm"

End Sub

Private Sub Comando1438_Click()
'Call AgregarTitulo  'Llamos al módulo para que nos ponga el título e icono al arrancar la aplicación.
'Me.sistemamusica = 1
'Call starplayer_Click
If Me.aloca = 0 Then
    MsgBox "Usuario No Atutorizado"
    Exit Sub
End If
'test para verificar conexion
If APIDisponible() Then
    DoCmd.OpenForm "LogueoUsuario"
    [Forms]![LogueoUsuario]![origenlogueo] = "[Introduccion]"
    'DoCmd.Close acForm, "Introduccion"

Else
    DoCmd.OpenForm "LogueoUsuarioAntiguo"
    [Forms]![LogueoUsuarioAntiguo]![origenlogueo] = "[Introduccion]"
    'DoCmd.Close acForm, "Introduccion"
End If



End Sub






Private Sub Comando1473_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmain"
End Sub

Private Sub Comando1474_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizardatosclub"
End Sub

Private Sub Comando1478_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablaslocales"

End Sub



Private Sub Form_Close()
If EsSistemaDeTienda() Then
    Call DetenerPingAutomatico
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim loca As Integer
Dim indi As String


Select Case CurrentProject.Name

    Case "ModuloPreDespacho.accdb", "ModuloPreDespacho - Copy.accdb", "ModuloPreDespacho - copia.accdb"
    
        DoCmd.OpenForm "MenuModuloPreDespacho"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloCompras.accdb", "ModuloCompras - Copy.accdb", "ModuloCompras - copia.accdb"
    
        DoCmd.OpenForm "Menu RegistroComprasIngresos"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloProduccion.accdb", "ModuloProduccion - Copy.accdb", "ModuloProduccion - copia.accdb"
        
        DoCmd.OpenForm "MenuPitayaProduccion0"
        DoCmd.Close acForm, "Introduccion"

    Case "ModuloContabilidad.accdb", "ModuloContabilidad - Copy.accdb", "ModuloContabilidad - copia.accdb"
        
        DoCmd.OpenForm "Menu Modulo Contabilidad"
        DoCmd.Close acForm, "Introduccion"
    
    Case "ModuloRRHH.accdb", "ModuloRRHH - Copy.accdb", "ModuloRRHH - copia.accdb"
        
        DoCmd.OpenForm "MenuModuloRRHH"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloAtencionAlCliente.accdb", "ModuloAtencionAlCliente - Copy.accdb", "ModuloAtencionAlCliente - copia.accdb"
        
        DoCmd.OpenForm "MenuModuloAtencionAlCliente"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloDespacho.accdb", "ModuloDespacho - Copy.accdb", "ModuloDespacho - copia.accdb"
        
        DoCmd.OpenForm "MenuModuloDespacho"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloAlmacen.accdb", "ModuloAlmacen - Copy.accdb", "ModuloAlmacen - copia.accdb"
        
        DoCmd.OpenForm "MenuModuloAlmacen"
        DoCmd.Close acForm, "Introduccion"
        
    Case "ModuloSupervisionSucursales.accdb", "ModuloSupervisionSucursales - Copy.accdb", "ModuloSupervisionSucursales - copia.accdb"
        
        DoCmd.OpenForm "MenuModuloSupervisionSucursales"
        DoCmd.Close acForm, "Introduccion"
        
    Case "Pitaya_System.accdb" 'sistema raiz
    
        Form.Caption = "BATIDOS PITAYA"
        Me.ShortcutMenu = False
        Me.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Background\Main.jpg"
        'Me.sistemamusica = 0
        
        'If Not indi = "_" Then
        '    Call EnviarMensajeTelegram("Abriendo Sistema",grupotoperaciones())
        'End If
        DoCmd.Restore
        Me.Requery
    
        [Forms]![Introduccion].Comando1437.Enabled = True 'Gestion
        [Forms]![Introduccion].Comando1438.Enabled = True 'Ventas
        [Forms]![Introduccion].Comando1217.Enabled = True 'Gestion
        [Forms]![Introduccion].senuelo.Visible = False
        DoCmd.Maximize
        
        Sleep 3000
        
        If CurrentProject.AllForms("MenuPitayaLogistica0").IsLoaded Or CurrentProject.AllForms("MenuPitayaProduccion0").IsLoaded Or CurrentProject.AllForms("Menu Gestion").IsLoaded Or CurrentProject.AllForms("Main Pitaya").IsLoaded Then
            'Ya no carga nuevamente el sistema
            Me.aloca = codigoLocal()
        Else
            'DoCmd.OpenForm "EleccionSucursalVigente"
            '[Forms]![EleccionSucursalVigente].tipocarga = 1
        End If

    Case Else ' sistema de algun local
        If CurrentProject.Name Like "Pitaya*" Then
            indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
            Me.aindi = indi
            
            'sistema de un local especifico
            Me.aloca = indi
        
            Form.Caption = "BATIDOS PITAYA"
            Me.ShortcutMenu = False
            Me.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Background\Main.jpg"
            'Me.sistemamusica = 0
            
            'If Not indi = "_" Then
            '    Call EnviarMensajeTelegram("Abriendo Sistema",grupotoperaciones())
            'End If
            DoCmd.Restore
            Me.Requery
            
            
            Sleep 5000
            
            [Forms]![Introduccion].Comando1437.Enabled = True 'Gestion
            [Forms]![Introduccion].Comando1438.Enabled = True 'Ventas
            [Forms]![Introduccion].Comando1217.Enabled = True 'Gestion

            [Forms]![Introduccion].senuelo.Visible = False
            DoCmd.Maximize
        Else
            MsgBox "Sistema actual no aplica"
        End If
End Select

If EsSistemaDeTienda() Then
    Call IniciarPingAutomatico
End If
End Sub




'Private Sub Form_Timer()

'If Me.sistemamusica <> 0 Then
'    If automusica.playState = 0 Or automusica.playState = 1 Then
'        Call starplayer_Click
'    End If
'End If
'End Sub





Private Sub Form_Timer()
If EsSistemaDeTienda() Then
    Call PingTimerTick
End If
End Sub
