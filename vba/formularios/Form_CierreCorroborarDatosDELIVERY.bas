' ==========================================================
' Modulo  : Form_CierreCorroborarDatosDELIVERY
' Tipo    : 100  |  Lineas: 56
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database
Public Sub forzaringresopos()

    DoCmd.OpenForm "CierreCorroborarDatosTRANSFERENCIAS"
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aCodOperario] = [Forms]![CierreCorroborarDatosDELIVERY]![aCodOperario]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aHoraFinal] = [Forms]![CierreCorroborarDatosDELIVERY]![aHoraFinal]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![afecha] = [Forms]![CierreCorroborarDatosDELIVERY]![afecha]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aTotalTransferenciaSistema] = [Forms]![Cierre por Turno]![transferenciasistema]
    
    [Forms]![Cierre por Turno]![aTotalPedidosYa] = [Forms]![CierreCorroborarDatosDELIVERY]![aTotalPedidosYa]
    DoCmd.Close acForm, "CierreCorroborarDatosDELIVERY"

End Sub





Private Sub Comando419_Click()
If absolutoabsoluto(Me.aTotalPedidosYa - Me.aTotalPedidosYaSistema) > 1 Then
    MsgBox "Verificar historial de ventas facturadas con PedidosYa y contrastar con el equipo de PedidosYa, solicitar desbloquear pedidos que requieran cambios"
Else
    DoCmd.OpenForm "CierreCorroborarDatosTRANSFERENCIAS"
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aCodOperario] = [Forms]![CierreCorroborarDatosDELIVERY]![aCodOperario]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aHoraFinal] = [Forms]![CierreCorroborarDatosDELIVERY]![aHoraFinal]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![afecha] = [Forms]![CierreCorroborarDatosDELIVERY]![afecha]
    [Forms]![CierreCorroborarDatosTRANSFERENCIAS]![aTotalTransferenciaSistema] = [Forms]![Cierre por Turno]![transferenciasistema]
    
    [Forms]![Cierre por Turno]![aTotalPedidosYa] = [Forms]![CierreCorroborarDatosDELIVERY]![aTotalPedidosYa]
    DoCmd.Close acForm, "CierreCorroborarDatosDELIVERY"
End If
End Sub

Private Sub Comando426_Click()
DoCmd.Close acForm, "Cierre por Turno"
DoCmd.Close acForm, "CierreCorroborarDatosDELIVERY"
End Sub

Private Sub Comando500_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]<>0 AND [Transferencia]=0"

End Sub

Private Sub Comando539_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "CierreCorroborarDatosDELIVERY"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



End Sub

