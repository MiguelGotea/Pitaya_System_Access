' ==========================================================
' Modulo  : Report_SoloQRpaquete
' Tipo    : 100
' Lineas  : 25
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub Comando212_Click()
DoCmd.PrintOut
End Sub

Private Sub Report_Open(Cancel As Integer)
'1 cm=567 twips
Me.Printer.TopMargin = 550
Me.Printer.BottomMargin = 0
Me.Printer.LeftMargin = 550
Me.Printer.RightMargin = 0

Dim archivo, codificacion As String

codificacion = "P" & [Forms]![PaquetesDePorciones]![porsub] * 10000 + [Forms]![PaquetesDePorciones]![cantidadxpaquete]
archivo = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\QR-Insumos\" & codificacion & ".png"
If Dir(archivo) <> "" Then
    Me.qrcode.Picture = archivo
Else
    Call GetQRCode(codificacion, 80, 80)
    Me.qrcode.Picture = archivo
End If

End Sub
