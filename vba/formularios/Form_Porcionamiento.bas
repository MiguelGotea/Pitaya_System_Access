' ==========================================================
' Modulo  : Form_Porcionamiento
' Tipo    : 100
' Lineas  : 17
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database






Private Sub Form_Close()
If CurrentProject.AllForms("Porcionamiento de Insumos").IsLoaded Then
    [Forms]![Porcionamiento de Insumos].Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
