' ==========================================================
' Modulo  : Form_CierreCorroborarDatosTRANSFERENCIAS
' Tipo    : 100  |  Lineas: 52
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Public Sub forzaringresopos()

    [Forms]![Cierre por Turno]![aTotalTransferencia] = [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aTotalTransferencia]
    
    Call Forms("Cierre por Turno").Comando419_Click
    DoCmd.Close acForm, "CierreCorroborarDatosTRANSFERENCIAS"

End Sub




Private Sub Comando419_Click()
If absolutoabsoluto(Me.aTotalTransferencia - Me.aTotalTransferenciaSistema) > 1 Then
    MsgBox "Verificar historial de ventas facturadas con transferencia y contrastar con los pagos por transferencia reportados"
Else

    [Forms]![Cierre por Turno]![aTotalTransferencia] = [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aTotalTransferencia]
    
    Call Forms("Cierre por Turno").Comando419_Click
    DoCmd.Close acForm, "CierreCorroborarDatosTRANSFERENCIAS"
End If



End Sub

Private Sub Comando426_Click()
DoCmd.Close acForm, "Cierre por Turno"
DoCmd.Close acForm, "CierreCorroborarDatosTRANSFERENCIAS"
End Sub

Private Sub Comando500_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]=0 AND [Transferencia]<>0"

End Sub

Private Sub Comando539_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "CierreCorroborarDatosTRANSFERENCIAS"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



End Sub

