' ==========================================================
' Modulo  : Form_BoletaExclusivoCentral
' Tipo    : 100
' Lineas  : 31
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database



Private Sub Comando426_Click()


Me.Comando426.height = 180
Me.Comando426.BackColor = RGB(255, 255, 255)
Me.Comando426.TopPadding = 0
Me.Comando426.LeftPadding = 0
Me.Comando426.RightPadding = 0
Me.Comando426.BottomPadding = 0
Sleep 1000
MsgBox "Factura guardado en el portapapeles"
Call CapturarVentanaComoPNG("C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Delivery Central\Facturas\" & Me.CodPedido & ".png")
Sleep 1000
Call CapturarFormularioAlPortapapeles("BoletaExclusivoCentral")
Sleep 1000
Me.Comando426.height = 360
Me.Comando426.BackColor = RGB(0, 0, 0)
Me.Comando426.TopPadding = 0
Me.Comando426.LeftPadding = 0
Me.Comando426.RightPadding = 0
Me.Comando426.BottomPadding = 0
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
