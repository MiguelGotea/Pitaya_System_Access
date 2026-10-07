' ==========================================================
' Modulo  : Form_DesempenoOperarioPorcionamiento
' Tipo    : 100
' Lineas  : 21
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database



Private Sub Comando175_Click()
Me.Requery
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

If CurrentProject.AllForms("Ingreso de Compras").IsLoaded Then
fechaac.Value = [Forms]![Ingreso de Compras]![fechaac]
End If

Me.Requery
End Sub
