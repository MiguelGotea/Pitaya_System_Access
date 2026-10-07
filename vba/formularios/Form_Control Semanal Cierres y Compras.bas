' ==========================================================
' Modulo  : Form_Control Semanal Cierres y Compras
' Tipo    : 100
' Lineas  : 59
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database

Private Sub Comando312_Click()
Me.Requery
Me.Filter = "semana = " & Me.asemana
Me.FilterOn = True

End Sub



Public Sub Comando496_Click()
Me.Requery
Me.Filter = "[semana] = " & Me.asemana
Me.FilterOn = True

Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Reporte Cierre Total.pdf"
DoCmd.OutputTo acOutputForm, "Control Semanal Cierres y Compras", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "Control Semanal Cierres y Compras"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "Control Semanal Cierres y Compras"
End Sub


Private Sub Fecha_DblClick(Cancel As Integer)
DoCmd.OpenForm "Control Diario Cierres y Compras"
[Forms]![Control Diario Cierres y Compras].dfecha = Me.Fecha
[Forms]![Control Diario Cierres y Compras].Form.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Filter = "Fecha= #" & Date & "#"
Me.FilterOn = True
Me.Requery
End Sub
