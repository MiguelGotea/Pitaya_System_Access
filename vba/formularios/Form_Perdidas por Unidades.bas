' ==========================================================
' Modulo  : Form_Perdidas por Unidades
' Tipo    : 100
' Lineas  : 63
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Private Sub Comando100_Click()
Me.Requery
End Sub



Private Sub Comando103_Click()
Me.afecha.Value = Me.afecha.Value - 1
Me.Requery
End Sub

Private Sub Comando128_Click()
DoCmd.OpenForm "RegistrarInsumosFijosConsumibles"
[Forms]![RegistrarInsumosFijosConsumibles]![fechaprocedencia] = Me.afecha
[Forms]![RegistrarInsumosFijosConsumibles]![adesdetabla] = "[Merma Cotizacion]"
[Forms]![RegistrarInsumosFijosConsumibles]![adesdeform] = "[Perdidas por Unidades]"
[Forms]![RegistrarInsumosFijosConsumibles]![guardado].width = 0
End Sub

Private Sub Comando139_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado", acNormal
[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.afecha
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[Merma Cotizacion]"
[Forms]![IngresoAutomaticoProductosFiltrado]![adesdeform] = "[Perdidas por Unidades]"

End Sub

Private Sub Comando217_Click()
DoCmd.OpenForm "RegistrarProductosPorciones"
[Forms]![RegistrarProductosPorciones]![fechaprocedencia] = Me.afecha
[Forms]![RegistrarProductosPorciones]![adesdetabla] = "[Merma Cotizacion]"
[Forms]![RegistrarProductosPorciones]![adesdeform] = "[Perdidas por Unidades]"
[Forms]![RegistrarProductosPorciones]![guardado].width = 0
End Sub

Private Sub Comando241_Click()
DoCmd.OpenForm "RegistrarProductosMostrador"
[Forms]![RegistrarProductosMostrador]![fechaprocedencia] = Me.afecha
[Forms]![RegistrarProductosMostrador]![adesdetabla] = "[Merma Cotizacion]"
[Forms]![RegistrarProductosMostrador]![adesdeform] = "[Perdidas por Unidades]"
[Forms]![RegistrarProductosMostrador]![guardado].width = 0
End Sub

Private Sub Comando257_Click()
DoCmd.OpenForm "RegistrarProductosNoPorciones"
[Forms]![RegistrarProductosNoPorciones]![fechaprocedencia] = Me.afecha
[Forms]![RegistrarProductosNoPorciones]![adesdetabla] = "[Merma Cotizacion]"
[Forms]![RegistrarProductosNoPorciones]![adesdeform] = "[Perdidas por Unidades]"
[Forms]![RegistrarProductosNoPorciones]![guardado].width = 0
End Sub


Private Sub Comando92_Click()
Me.afecha.Value = Me.afecha.Value + 1
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
