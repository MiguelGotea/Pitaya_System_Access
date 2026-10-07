' ==========================================================
' Modulo  : Form_HistorialCostoUnitarioIngredientes
' Tipo    : 100  |  Lineas: 54
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database



Public Sub Comando206_Click()
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

If codigoLocal() = 0 Then
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Reporte Semanal\" & Me.asemana
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Reporte Semanal"
Else
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
End If

If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If


archivo = Direccion & "\" & "Historial de Costos Unitarios.pdf"
DoCmd.OutputTo acOutputForm, "HistorialCostoUnitarioIngredientes", acFormatPDF, archivo, False

DoCmd.Close acForm, "HistorialCostoUnitarioIngredientes"

End Sub

Private Sub Comando52_Click()
Me.Requery

End Sub






Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaant = numerosemana(Date)
End Sub
