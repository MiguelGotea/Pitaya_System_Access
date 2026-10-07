' ==========================================================
' Modulo  : Form_DetalleCompraxProovedor
' Tipo    : 100
' Lineas  : 40
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database

Public Sub Comando324_Click()
Me.Requery
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim Direccion3 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.asemana & "\Compras Quincenales"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.asemana
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras"

If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Pedido Proovedor " & DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.cproov & "") & ".pdf"
DoCmd.OutputTo acOutputForm, "DetalleCompraxProovedor", acFormatPDF, archivo, False

DoCmd.Close acForm, "DetalleCompraxProovedor"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
