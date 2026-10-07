' ==========================================================
' Modulo  : Report_Boleta
' Tipo    : 100  |  Lineas: 14
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database




Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 550
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 550
Me.Printer.RightMargin = 0
Me.imagen.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Logo\Logo.png"

End Sub
