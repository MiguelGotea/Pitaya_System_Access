' ==========================================================
' Modulo  : Report_Compras Fijos
' Tipo    : 100
' Lineas  : 9
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
End Sub
