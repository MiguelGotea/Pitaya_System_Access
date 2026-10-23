' ==========================================================
' Modulo  : Report_KardexIngredienteSemana
' Tipo    : 100  |  Lineas: 9
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:04
' ==========================================================

Option Compare Database

Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 1440
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 567
Me.Printer.RightMargin = 0
End Sub
