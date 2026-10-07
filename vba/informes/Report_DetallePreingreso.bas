' ==========================================================
' Modulo  : Report_DetallePreingreso
' Tipo    : 100  |  Lineas: 11
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database


Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 550
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 550
Me.Printer.RightMargin = 0

End Sub
