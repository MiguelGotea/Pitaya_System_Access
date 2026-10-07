' ==========================================================
' Modulo  : Form_Control Existencias Pareto
' Tipo    : 100
' Lineas  : 46
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Public Sub Comando119_Click()
Me.Requery
Me.Filter = "Posicion <=25 and Posicion >0"
Me.FilterOn = True
Me.OrderBy = "Posicion ASC"
Me.OrderByOn = True
Me.vari = FRVariacionTotal(Me.asemana) 'FRVariacion
Me.merm = ValorizacionMermasSemana(Me.asemana)

End Sub

Public Sub Comando437_Click()
If codigoLocal() = 0 Then 'Sistema local
    Exit Sub
End If

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

archivo = Direccion & "\" & Me.asemana & " - 2.Control Existencias Pareto.pdf"
DoCmd.OutputTo acOutputForm, "Control Existencias Pareto", acFormatPDF, archivo, False
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
Me.OrderBy = ""
Me.OrderByOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
