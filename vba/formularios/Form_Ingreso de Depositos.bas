' ==========================================================
' Modulo  : Form_Ingreso de Depositos
' Tipo    : 100
' Lineas  : 68
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando50_Click()
Me.Requery

End Sub


Private Sub Comando56_Click()
Dim Direccion1, Direccion2, Direccion3 As String
Dim Mes, ano As Long
Mes = Me.mesac
ano = Me.añoac
Direccion1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal() & "\" & ano & "\" & Mes
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal() & "\" & ano
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal()

If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion1, vbDirectory) = "" Then
    MkDir Direccion1
End If

Dim archivo As String
archivo = Direccion1 & "\Depositos Totales del Mes.pdf"
DoCmd.OutputTo acOutputForm, "Ingreso de Depositos", acFormatPDF, archivo, False

End Sub

Private Sub Comando68_Click()

If mesac.Value = 1 Then
    mesac.Value = 12
    añoac.Value = añoac.Value - 1
    Me.Requery
Else
    mesac.Value = mesac.Value - 1
    Me.Requery
End If
End Sub

Private Sub Comando71_Click()

If mesac.Value = 12 Then
    mesac.Value = 1
    añoac.Value = añoac.Value + 1
    Me.Requery
Else
    mesac.Value = mesac.Value + 1
    Me.Requery
End If
End Sub

Private Sub Form_Load()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
