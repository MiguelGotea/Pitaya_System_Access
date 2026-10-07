' ==========================================================
' Modulo  : Form_CierreCorroborarDatosPOS
' Tipo    : 100
' Lineas  : 55
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Public Sub forzaringresopos()

    DoCmd.OpenForm "CierreCorroborarDatosDELIVERY"
    [Forms]![CierreCorroborarDatosDELIVERY]![aCodOperario] = [Forms]![CierreCorroborarDatosPOS]![aCodOperario]
    [Forms]![CierreCorroborarDatosDELIVERY]![aHoraFinal] = [Forms]![CierreCorroborarDatosPOS]![aHoraFinal]
    [Forms]![CierreCorroborarDatosDELIVERY]![afecha] = [Forms]![CierreCorroborarDatosPOS]![afecha]
    [Forms]![CierreCorroborarDatosDELIVERY]![aTotalPedidosYaSistema] = [Forms]![Cierre por Turno]![deliverysistema]
    
    [Forms]![Cierre por Turno]![aTotalPOS] = [Forms]![CierreCorroborarDatosPOS]![aTotalPOS]
    DoCmd.Close acForm, "CierreCorroborarDatosPOS"

End Sub




Private Sub Comando419_Click()
If absolutoabsoluto(Me.aTotalPOS - Me.aTotalPOSSistema) > 1 Then
    MsgBox "Verificar historial de ventas facturadas con POS y contrastar con el POS del banco, solicitar desbloquear pedidos que requieran cambios"
Else
    DoCmd.OpenForm "CierreCorroborarDatosDELIVERY"
    [Forms]![CierreCorroborarDatosDELIVERY]![aCodOperario] = [Forms]![CierreCorroborarDatosPOS]![aCodOperario]
    [Forms]![CierreCorroborarDatosDELIVERY]![aHoraFinal] = [Forms]![CierreCorroborarDatosPOS]![aHoraFinal]
    [Forms]![CierreCorroborarDatosDELIVERY]![afecha] = [Forms]![CierreCorroborarDatosPOS]![afecha]
    [Forms]![CierreCorroborarDatosDELIVERY]![aTotalPedidosYaSistema] = [Forms]![Cierre por Turno]![deliverysistema]
    
    [Forms]![Cierre por Turno]![aTotalPOS] = [Forms]![CierreCorroborarDatosPOS]![aTotalPOS]
    DoCmd.Close acForm, "CierreCorroborarDatosPOS"
End If
End Sub

Private Sub Comando426_Click()
DoCmd.Close acForm, "Cierre por Turno"
DoCmd.Close acForm, "CierreCorroborarDatosPOS"
End Sub

Private Sub Comando500_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]=0 AND [Transferencia]=0"
End Sub

Private Sub Comando539_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "CierreCorroborarDatosPOS"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



End Sub

