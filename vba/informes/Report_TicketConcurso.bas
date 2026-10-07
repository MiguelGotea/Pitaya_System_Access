' ==========================================================
' Modulo  : Report_TicketConcurso
' Tipo    : 100  |  Lineas: 11
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database

Private Sub Report_Open(Cancel As Integer)
Me.concurso.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Ticket Concurso\Concurso.jpeg"

'1 cm=567 twips
Me.Printer.TopMargin = 550
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 550
Me.Printer.RightMargin = 0
End Sub
