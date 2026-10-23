' ==========================================================
' Modulo  : Form_HistorialDespachoInsumosFIjos
' Tipo    : 100  |  Lineas: 18
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:05
' ==========================================================

Option Compare Database

Private Sub Comando52_Click()
Me.Requery

End Sub






Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

Me.semaactual = numerosemana(Date)

End Sub
