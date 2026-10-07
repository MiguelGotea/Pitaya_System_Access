' ==========================================================
' Modulo  : Report_StickerEquiposCodigo
' Tipo    : 100  |  Lineas: 13
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Private Sub Comando212_Click()
DoCmd.PrintOut
End Sub

Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
End Sub
