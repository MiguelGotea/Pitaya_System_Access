' ==========================================================
' Modulo  : Form_Control Existencias Compra Venta
' Tipo    : 100
' Lineas  : 55
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Private Sub Comando54_Click()
Me.Requery
End Sub

Public Sub Comando119_Click()
Me.Requery
End Sub

Public Sub Comando352_Click()
If codigoLocal() = 0 Then 'Sistema local
    Exit Sub
End If

Me.Requery
Me.Printer.Orientation = acPRORPortrait
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

archivo = Direccion & "\" & "Control de Productos Mostrador.pdf"
DoCmd.OutputTo acOutputForm, "Control Existencias Compra Venta", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "Control Existencias Compra Venta"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "Control Existencias Compra Venta"
End Sub

Private Sub Form_Close()
Me.OrderBy = ""
Me.OrderByOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
