' ==========================================================
' Modulo  : Form_Control Procesamiento / Ingresos Global
' Tipo    : 100
' Lineas  : 52
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
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
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Reporte Semanal\" & Me.asemana
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Reporte Semanal"
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

archivo = Direccion & "\" & "Procesamiento vs Ingresos.pdf"
DoCmd.OutputTo acOutputForm, "Control Procesamiento / Ingresos Global", acFormatPDF, archivo, False

DoCmd.Close acForm, "Control Procesamiento / Ingresos Global"
End Sub

Private Sub Comando52_Click()
Me.Requery

End Sub






Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaant = numerosemana(Date)
End Sub
