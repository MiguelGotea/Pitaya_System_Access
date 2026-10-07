' ==========================================================
' Modulo  : Form_Control Mermas/Consumos
' Tipo    : 100
' Lineas  : 19
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando54_Click()
Me.Requery
End Sub

Private Sub Comando98_Click()
On Error GoTo Error
DoCmd.OutputTo acOutputForm, "Control Mermas/Consumos", acFormatXLSX, "C:\Users\" & NombreSistema() & "\Desktop\Sistema\" & nombrelocal() & ".xlsx", False
MsgBox "Datos Exportados"
Exit Sub
Error:
MsgBox "Ingresar Datos Correctamente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
