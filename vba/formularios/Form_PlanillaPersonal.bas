' ==========================================================
' Modulo  : Form_PlanillaPersonal
' Tipo    : 100
' Lineas  : 86
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Private Sub Comando198_Click()
Me.Requery
End Sub


Private Sub Comando403_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String
Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If
archivo = Direccionmes & "\Boleta Final " & Me.intervaloboleta & " - " & Sucursal & " - " & NombreOperario(Me.CodOperario) & " " & ".pdf"

DoCmd.OpenReport "BoletaOperario", acViewReport, , "[CodBoletaPagoOperario]=" & [Forms]![PlanillaPersonal]![CodBoletaPagoOperario]

[Reports]![BoletaOperario].Report.Printer.Orientation = acPRORLandscape
[Reports]![BoletaOperario].Report.Printer.LeftMargin = 0
[Reports]![BoletaOperario].Report.Printer.RightMargin = 0
[Reports]![BoletaOperario].Report.Printer.BottomMargin = 0
[Reports]![BoletaOperario].Report.Printer.TopMargin = 0


DoCmd.OutputTo acOutputReport, "BoletaOperario", acFormatPDF, archivo, False
DoCmd.Close acReport, "BoletaOperario"
End Sub

Private Sub Comando754_Click()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim cot As Integer
Dim Cal As Double

miSQL = "SELECT Operarios.Sucursal, Operarios.Operativo, Operarios.CodOperario FROM Operarios" & _
" WHERE (((Operarios.Sucursal)=" & Me.Sucursal & ") AND ((Operarios.Operativo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO BoletaPagoOperario(CodOperario, Quincena) values" & _
    " (" & rst("CodOperario") & ", #" & Me.desde & "#)"
    DoCmd.SetWarnings True

    rst.MoveNext
Next I

rst.Close
MsgBox "Planilla creada correctamente de la quincena"
Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se pudo generar planilla correctamente"
End Sub



Private Sub Comando934_Click()
DoCmd.OpenForm "IngresoBonosMensuales"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

Private Sub intervalo_Change()
Me.desde.Requery
Me.hasta.Requery

End Sub


