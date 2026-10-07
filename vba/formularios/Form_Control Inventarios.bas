' ==========================================================
' Modulo  : Form_Control Inventarios
' Tipo    : 100
' Lineas  : 58
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database



Private Sub Comando458_Click()
Me.Requery
End Sub

Private Sub Comando507_Click()
Me.semana1 = Me.semana1 - 1
Me.Requery
End Sub

Private Sub Comando511_Click()
Me.semana1 = Me.semana1 + 1
Me.Requery
End Sub



Public Sub Comando558_Click()
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

archivo = Direccion & "\" & "Control de Inventario.pdf"
DoCmd.OutputTo acOutputForm, "Control Inventarios", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "Control Inventarios"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "Control Inventarios"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
