' ==========================================================
' Modulo  : Form_HistoricoConsumoPorcionesSUcursales
' Tipo    : 100
' Lineas  : 42
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database


Public Sub Comando170_Click()
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

DoCmd.SelectObject acForm, "HistoricoConsumoPorcionesSucursales", True
'DoCmd.RunCommand acCmdPrint

Dim Direccion1 As String
Dim archivo As String

Direccion1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Historico Consumos\" & Me.semanaactual

If Dir(Direccion1, vbDirectory) = "" Then
    MkDir Direccion1
End If

archivo = Direccion1 & "\" & Me.asucursal & " - Historico de Consumo Porciones.pdf"
DoCmd.OutputTo acOutputForm, "HistoricoConsumoPorcionesSucursales", acFormatPDF, archivo, False

DoCmd.Close acForm, "HistoricoConsumoPorcionesSucursales"
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
