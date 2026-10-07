' ==========================================================
' Modulo  : Form_EstandarInsumosFijosConsumiblesGlobal
' Tipo    : 100
' Lineas  : 54
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database





Public Sub Comando199_Click()
Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Me.Requery

Dim Direccion, Direccion2 As String
Dim Direccion3 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.semanaactual & "\Compras Quincenales"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.semanaactual
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

archivo = Direccion & "\" & "Estandar Insumos Fijos Consumibles.pdf"
DoCmd.OutputTo acOutputForm, "EstandarInsumosFijosConsumiblesGlobal", acFormatPDF, archivo, False

DoCmd.Close acForm, "EstandarInsumosFijosConsumiblesGlobal"
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub


Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub
