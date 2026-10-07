' ==========================================================
' Modulo  : Form_Ingresos a Pitaya
' Tipo    : 100  |  Lineas: 59
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database

Private Sub Comando106_Click()
DoCmd.OpenForm "IngresoDePorciones", acNormal
[Forms]![IngresoDePorciones]![fechaprocedencia] = Me.fechaac
End Sub

Private Sub Comando108_Click()
DoCmd.OpenForm "RegistrarProductoMarcaPitaya"
[Forms]![RegistrarProductoMarcaPitaya]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductoMarcaPitaya]![adesde] = "[IngresosPitaya]"
End Sub

Private Sub Comando112_Click()
fechaac.Value = fechaac.Value + 1
Me.Requery
End Sub

Private Sub Comando113_Click()
fechaac.Value = fechaac.Value - 1
Me.Requery
End Sub

Private Sub Comando139_Click()
DoCmd.OpenForm "IngresoAutomaticoProductos", acNormal
[Forms]![IngresoAutomaticoProductos]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductos]![adestino] = "[IngresosPitaya]"
[Forms]![IngresoAutomaticoProductos]![adesdeform] = "[Ingresos a Pitaya]"
End Sub

Private Sub Comando148_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If
End Sub



Private Sub Comando175_Click()
DoCmd.OpenForm "IngresoCodigoBarras"
[Forms]![IngresoCodigoBarras]![fechaprocedencia] = Me.fechaac
[Forms]![IngresoCodigoBarras]![adesde] = "[IngresosPitaya]"
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

If CurrentProject.AllForms("Ingreso de Compras").IsLoaded Then
fechaac.Value = [Forms]![Ingreso de Compras]![fechaac]
End If

Me.Requery
End Sub
