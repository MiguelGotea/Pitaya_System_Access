' ==========================================================
' Modulo  : Form_Menu PITAYA Global
' Tipo    : 100  |  Lineas: 101
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================



Private Sub Comando5056_Click()
If codigoLocal() = 19 Then
    DoCmd.OpenForm "Menu Mostrador Unica"
Else
    DoCmd.OpenForm "Menu Mostrador"
End If
End Sub

Private Sub Comando538_Click()
On Error Resume Next
''Version cerradno y abir nuevamente
'Dim codi As Long
'Dim men As String
'men = Me.Name ' nombre del formulario
'codi = [Forms]![Nota de Pedido]![CodPedido]
'DoCmd.Close acForm, "Nota de Pedido"
'DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & codi
'DoCmd.Close acForm, men

'Version Fast
If Me.solovista = 0 Then
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Nota de Pedido].Form.Secundario57.Requery
    [Forms]![Nota de Pedido].Form.ppedido.Requery
    [Forms]![Nota de Pedido].Form.Texto193.Requery
    [Forms]![Nota de Pedido].Form.Texto1278.Requery

    [Forms]![Nota de Pedido].Form.cambiocor.Requery
    [Forms]![Nota de Pedido].Form.recibecor.SetFocus
End If
[Forms]![Nota de Pedido].Form.Secundario57.Requery
DoCmd.Close acForm, "Menu PITAYA Global"
End Sub



Private Sub Comando6965_Click()

DoCmd.OpenForm "MenuCombos"

End Sub

Private Sub Comando6971_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "Menu Descuento Especial"
End Sub

Private Sub Comando7178_Click()
On Error GoTo Nulo
Dim numeroclub As String
numeroclub = InputBox("Ingresar numero de membresia asignado", "Membresia")

If CInt(numeroclub) < 1000 Then
    MsgBox "Ingresar numero valido de MEMBRESIA, volver a facturar"
    Exit Sub
End If

If numeroclub = "" Or IsNull(numeroclub) Then
    MsgBox "ingresar un numero correcto de membresia para poder agregar membresia a la factura"
    Exit Sub
Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, Vinculo, Observaciones)" & _
    " VALUES ('TCPv4',1, " & [Forms]![Nota de Pedido]![CodPedido] & ", 0, '-', 0, '" & numeroclub & "')"
    DoCmd.SetWarnings True
End If


Exit Sub

Nulo:
MsgBox "ingresar un numero correcto de membresia para poder agregar membresia a la factura"

End Sub

Private Sub Comando7201_Click()
If codigoLocal() = 15 Then
    DoCmd.OpenForm "Menu Servicio Delivery Central"
ElseIf DLookup("[Ciudad]", "StatusSucursales", "[CodLocal]=" & codigoLocal()) = "Managua" Then
    DoCmd.OpenForm "Menu Servicio Delivery Managua"
Else
    MsgBox "No tiene acceso al menu de servicio delivery"
End If
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



End Sub

Private Sub salirsinguardar_Click()
DoCmd.Close
End Sub

