' ==========================================================
' Modulo  : Form_Costeo Inventario Cotizacion
' Tipo    : 100
' Lineas  : 23
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database

Private Sub Comando2_Click()
Me.total = FRCostoIFCotizacion(Me.asemana)
Me.Requery
End Sub

Private Sub Comando79_Click()
Me.asemana = Me.asemana - 1
Me.total = FRCostoIFCotizacion(Me.asemana)
Me.Requery
End Sub

Private Sub Comando82_Click()
Me.asemana = Me.asemana + 1
Me.total = FRCostoIFCotizacion(Me.asemana)
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
