' ==========================================================
' Modulo  : Form_Relacion Pedidos Anulados
' Tipo    : 100  |  Lineas: 40
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database

Public Sub Comando119_Click()
Me.Filter = "semana = " & Me.asemana
Me.FilterOn = True
Me.Requery
End Sub

Public Sub Comando437_Click()

If codigoLocal() = 0 Then 'Sistema local
    Exit Sub
End If

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.asemana & " - 7.Pedidos Anulados.pdf"
DoCmd.OutputTo acOutputForm, "Relacion Pedidos Anulados", acFormatPDF, archivo, False
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
