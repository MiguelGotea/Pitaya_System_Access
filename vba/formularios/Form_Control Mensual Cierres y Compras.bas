' ==========================================================
' Modulo  : Form_Control Mensual Cierres y Compras
' Tipo    : 100
' Lineas  : 57
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database



Private Sub Comando696_Click()
Me.Requery
Me.Filter = "mes = " & Me.ames & " AND ano= " & Me.aano
Me.FilterOn = True

End Sub

Private Sub Comando496_Click()
Me.Requery
Me.Filter = "mes = " & Me.ames & " AND ano= " & Me.aano
Me.FilterOn = True

Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccionano, Direccionmes As String
Dim archivo As String
Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.aano
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.aano & "\" & Me.ames
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

archivo = Direccionmes & "\" & Me.ames & "_" & Me.aano & " - " & codigoLocal() & " - Reporte Cierre Mensual.pdf"
DoCmd.OutputTo acOutputForm, "Control Mensual Cierres y Compras", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "Control Mensual Cierres y Compras"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "Control Mensual Cierres y Compras"
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
