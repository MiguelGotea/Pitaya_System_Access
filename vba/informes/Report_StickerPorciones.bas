' ==========================================================
' Modulo  : Report_StickerPorciones
' Tipo    : 100
' Lineas  : 23
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database

Private Sub Comando212_Click()
DoCmd.PrintOut
End Sub

Private Sub Report_Open(Cancel As Integer)
Dim archivo, codificacion As String

If [Forms]![PaquetesDePorciones]![especiales] Like "ESPECIAL*" Then
    codificacion = [Forms]![PaquetesDePorciones]![especiales]
Else
    codificacion = "P" & [Forms]![PaquetesDePorciones]![porsub] * 10000 + [Forms]![PaquetesDePorciones]![cantidadxpaquete]
End If

archivo = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\QR-Insumos\" & codificacion & ".png"
If Dir(archivo) <> "" Then
    Me.qrcode.Picture = archivo
Else
    Call GetQRCode(codificacion, 80, 80)
    Me.qrcode.Picture = archivo
End If
End Sub
