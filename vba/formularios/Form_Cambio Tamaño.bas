' ==========================================================
' Modulo  : Form_Cambio Tamaño
' Tipo    : 100  |  Lineas: 26
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando7_Click()
DoCmd.Close acForm, "Cambio Tamaño", acSaveYes
[Forms]![Nota de Pedido].Form.Requery
[Forms]![Nota de Pedido].Form.Secundario57.Requery
[Forms]![Nota de Pedido].Form.Texto1278.Requery

[Forms]![Nota de Pedido].Form.recibecor.SetFocus
End Sub

Private Sub Form_Close()
If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Nota de Pedido].Form.Secundario57.Requery
    [Forms]![Nota de Pedido].Form.Texto1278.Requery

    [Forms]![Nota de Pedido].Form.recibecor.SetFocus
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

