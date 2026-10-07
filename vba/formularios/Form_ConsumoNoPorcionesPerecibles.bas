' ==========================================================
' Modulo  : Form_ConsumoNoPorcionesPerecibles
' Tipo    : 100
' Lineas  : 42
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database

Public Sub Comando102_Click()
Me.Requery
End Sub

Public Sub Comando170_Click()
Me.Requery
DoCmd.SelectObject acForm, "ConsumoNoPorcionesPerecibles", True
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String
Dim archivo2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & Me.asemana & "\Detalle Sucursales"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & Me.asemana
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.alocali & " - Productos Compras Locales.pdf"
DoCmd.OutputTo acOutputForm, "ConsumoNoPorcionesPerecibles", acFormatPDF, archivo, False

DoCmd.Close acForm, "ConsumoNoPorcionesPerecibles"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
End Sub
