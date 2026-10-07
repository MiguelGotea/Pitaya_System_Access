' ==========================================================
' Modulo  : Form_Musica Pitaya
' Tipo    : 100
' Lineas  : 73
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando20_Click()
Call starplayer_Click
End Sub

Private Sub Comando5_Click()
DoCmd.Minimize
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub Form_Timer()

If Me.cancion = "No Music" Then
DoCmd.Minimize
'DoCmd.OpenForm "Main Pitaya"
Call starplayer_Click
End If

End Sub

Private Sub stop_Click()
automusica.Object.Controls.pause
End Sub

Private Sub resume_Click()
automusica.Object.Controls.play
End Sub


Private Sub starplayer_Click()
Dim ruta As String
ruta = "C:\Users\" & NombreSistema() & "\Google Drive BP\Pitaya Music\"
Do While 1 = 1

Set fso = CreateObject("Scripting.FileSystemObject")
Set carpeta = fso.GetFolder(ruta)

Randomize
I = CInt((Rnd() * carpeta.Files.Count) + 1)
j = 1

For Each archivo In carpeta.Files

If j >= I Then
    
    If archivo.Type = "MP3 Format Sound" Or archivo.Type = "Sonido en formato MP3" Then
        automusica.url = ruta & archivo.Name
        cancion.Value = archivo.Name
        automusica.Object.Controls.play
        automusica.settings.volume = DLookup("[VolumenMusica]", "DatosSistema")

        Do Until automusica.playState = 1
        DoEvents
        Loop
    End If
    
End If
j = j + 1

Next archivo

Set fso = Nothing
Set carpeta = Nothing
Set ficheros = Nothing

Loop

End Sub
