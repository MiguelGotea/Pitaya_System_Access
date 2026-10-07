' ==========================================================
' Modulo  : Form_Control Depositos Cierres
' Tipo    : 100  |  Lineas: 60
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database



Private Sub Comando201_Click()
Me.afecha.Value = Me.afecha.Value - 1
End Sub

Private Sub Comando205_Click()
Me.dfecha.Value = Me.dfecha.Value + 1
End Sub

Private Sub Comando207_Click()
Me.afecha.Value = Me.afecha.Value + 1
End Sub

Private Sub Comando230_Click()
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion1, Direccion2, Direccion3 As String
Dim Mes, ano As Long
Mes = Month(Me.afecha)
ano = Year(Me.afecha)
Direccion1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal() & "\" & ano & "\" & Mes
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal() & "\" & ano
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Contabilidad\Control Depositos\" & nombrelocal()

If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion1
End If
If Dir(Direccion1, vbDirectory) = "" Then
    MkDir Direccion1
End If

Dim archivo As String
archivo = Direccion1 & "\Desde " & Day(Me.dfecha) & "(" & Month(Me.dfecha) & ")" & " Hasta " & Day(Me.afecha) & "(" & Month(Me.afecha) & ").pdf"
DoCmd.OutputTo acOutputForm, "Control Depositos Cierres", acFormatPDF, archivo, False
End Sub

Private Sub Comando33_Click()
Me.Requery
End Sub

Private Sub dmenos_Click()
Me.dfecha.Value = Me.dfecha.Value - 1
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
