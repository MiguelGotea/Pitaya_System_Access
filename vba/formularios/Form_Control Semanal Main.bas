' ==========================================================
' Modulo  : Form_Control Semanal Main
' Tipo    : 100
' Lineas  : 59
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database




Public Sub Comando157_Click()
On Error GoTo CerrarYa

Me.Requery
Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

If codigoLocal() = 0 Then
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal"
Else
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
End If

If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Control de No Porciones Desde " & Me.semanadc & " Hasta " & Me.semanaac & ".pdf"
DoCmd.OutputTo acOutputForm, "Control Semanal Main", acFormatPDF, archivo, False

DoCmd.Close acForm, "Control Semanal Main"
Exit Sub

CerrarYa:

DoCmd.Close acForm, "Control Semanal Main"
End Sub




Private Sub Comando219_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaac = numerosemana(Date) - 1
Me.semanadc = numerosemana(Date) - 1
Me.InsideHeight = 8000
End Sub
