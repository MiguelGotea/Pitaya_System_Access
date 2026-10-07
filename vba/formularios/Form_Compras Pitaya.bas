' ==========================================================
' Modulo  : Form_Compras Pitaya
' Tipo    : 100
' Lineas  : 36
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Public Sub Comando85_Click()
DoCmd.OpenReport "Compras_Quincenal", acViewReport

DoCmd.SelectObject acReport, "Compras_Quincenal", True
DoCmd.RunCommand acCmdPrint

Dim Direccion, Direccion2 As String
Dim archivo As String
Dim archivo2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras"
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Productos Generales" & " (" & Me.csemana & " semanas).pdf"
DoCmd.OutputTo acOutputReport, "Compras_Quincenal", acFormatPDF, archivo, False


DoCmd.Close acReport, "Compras_Quincenal"
DoCmd.Close acForm, "Compras Pitaya"

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaac = numerosemana(Date)
Me.ShortcutMenu = False
End Sub
