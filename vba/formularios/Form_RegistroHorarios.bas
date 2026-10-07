' ==========================================================
' Modulo  : Form_RegistroHorarios
' Tipo    : 100  |  Lineas: 56
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database








Private Sub Comando108_Click()
Me.Requery
End Sub

Private Sub Comando148_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If
Me.Requery
End Sub



Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Comando279_Click()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TarifaPersonal.CodTarifa, TarifaPersonal.CodOperario, TarifaPersonal.FechaInicial," & _
" TarifaPersonal.FechaFinal, TarifaPersonal.Tarifa," & _
" TarifaPersonal.HL, TarifaPersonal.HMA, TarifaPersonal.HMI, TarifaPersonal.HJ, TarifaPersonal.HV, TarifaPersonal.HS, TarifaPersonal.HD" & _
" FROM TarifaPersonal WHERE (((TarifaPersonal.CodTarifa)=" & Me.CodTarifa & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO TarifaPersonal(CodOperario, FechaInicial, FechaFinal, Tarifa, HL, HMA, HMI, HJ, HV, HS, HD) values" & _
" (" & rst("CodOperario") & ", #" & rst("FechaInicial") & "#, #" & rst("FechaFinal") & "#, " & rst("Tarifa") & _
", " & rst("HL") & ", " & rst("HMA") & ", " & rst("HMI") & ", " & rst("HJ") & ", " & rst("HV") & ", " & rst("HS") & ", " & rst("HD") & ")"
DoCmd.SetWarnings True

rst.Close
Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se pudo copiar el registro"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
