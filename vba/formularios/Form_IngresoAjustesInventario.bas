' ==========================================================
' Modulo  : Form_IngresoAjustesInventario
' Tipo    : 100
' Lineas  : 55
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database


Private Sub Comando155_Click()
DoCmd.OpenForm "RegistrarProductosPorciones"
[Forms]![RegistrarProductosPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosPorciones]![adesdetabla] = "[AjustesInventario]"
[Forms]![RegistrarProductosPorciones]![adesdeform] = "[IngresoAjustesInventario]"
End Sub

Private Sub Comando174_Click()
DoCmd.OpenForm "RegistrarProductosMostrador"
[Forms]![RegistrarProductosMostrador]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosMostrador]![adesdetabla] = "[AjustesInventario]"
[Forms]![RegistrarProductosMostrador]![adesdeform] = "[IngresoAjustesInventario]"
[Forms]![RegistrarProductosMostrador]![guardado].Visible = False
[Forms]![RegistrarProductosMostrador]![guardado].width = 0

End Sub

Private Sub Comando234_Click()
DoCmd.OpenForm "RegistrarProductosNoPorciones"
[Forms]![RegistrarProductosNoPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosNoPorciones]![adesdetabla] = "[AjustesInventario]"
[Forms]![RegistrarProductosNoPorciones]![adesdeform] = "[IngresoAjustesInventario]"
[Forms]![RegistrarProductosNoPorciones]![guardado].Visible = False
[Forms]![RegistrarProductosNoPorciones]![guardado].width = 0
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub

Private Sub Comando21_Click()
Me.fechaac = Me.fechaac - 1
Me.Requery
End Sub

Private Sub Comando23_Click()
Me.fechaac = Me.fechaac + 1
Me.Requery
End Sub



Private Sub fechaac_Change()
Me.Requery
End Sub






