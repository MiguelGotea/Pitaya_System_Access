' ==========================================================
' Modulo  : Form_KardexCotizacionSemanal
' Tipo    : 100
' Lineas  : 17
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database



Private Sub Comando184_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

If CurrentProject.AllForms("Ingreso de Compras").IsLoaded Then
fechaac.Value = [Forms]![Ingreso de Compras]![fechaac]
End If

Me.Requery
End Sub
