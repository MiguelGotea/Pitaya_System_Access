' ==========================================================
' Modulo  : Form_Gestion Objetos SIstema
' Tipo    : 100
' Lineas  : 17
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Private Sub Comando83_Click()
MsgBox "Tables_Local = 1" & vbCrLf & "Tables_Linked_ODBC = 4" & vbCrLf & "Tables_Linked = 6" & vbCrLf & "Queries = 5" & vbCrLf & "Forms = -32768" & _
vbCrLf & "Reports = -32764" & vbCrLf & "Macros = -32766" & vbCrLf & "Modules = -32761" & vbCrLf & "/ : DIreccion" & vbCrLf & "- : SUbformulario"
End Sub

Private Sub ctipo_Change()
Me.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
