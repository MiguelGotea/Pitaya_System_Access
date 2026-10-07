' ==========================================================
' Modulo  : Form_ProporcionadoIngredientes
' Tipo    : 100
' Lineas  : 11
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando14_Click()
Me.pesototal = InputBox("Ingrese Peso Total a Porcionar")
DoCmd.OpenReport "Medidas Ingredientes", acViewNormal
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
