' ==========================================================
' Modulo  : Form_RegistroEquipos
' Tipo    : 100
' Lineas  : 33
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database









Private Sub Comando148_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If
End Sub







Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

If CurrentProject.AllForms("Ingreso de Compras").IsLoaded Then
fechaac.Value = [Forms]![Ingreso de Compras]![fechaac]
End If

Me.Requery
End Sub
