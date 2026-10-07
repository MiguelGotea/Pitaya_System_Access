' ==========================================================
' Modulo  : Form_Control Insumos Importantes
' Tipo    : 100
' Lineas  : 62
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database



Private Sub Comando458_Click()
Me.Requery
End Sub





Public Sub Comando562_Click()
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.semana1
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Control de Insumos Importantes.pdf"
DoCmd.OutputTo acOutputForm, "Control Insumos Importantes", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "Control Insumos Importantes"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "Control Insumos Importantes"
End Sub



Private Sub Comando793_Click()
DoCmd.OpenForm "Ingreso Inventario Pitaya"
End Sub

Private Sub Comando813_Click()
DoCmd.OpenForm "Ingresos a Pitaya"
End Sub

Private Sub Comando815_Click()
DoCmd.OpenForm "ElegirTipoDeMerma"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
