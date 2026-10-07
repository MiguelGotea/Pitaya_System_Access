' ==========================================================
' Modulo  : Form_Menu Cupones Web
' Tipo    : 100  |  Lineas: 28
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:20
' ==========================================================



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

Private Sub regresar_Click()
If APIDisponible() Then
    Dim montocupondato As Integer
    Call DescargarTablaCompleta("cupones_sucursales", "cupones_sucursales", "")
    montocupondato = montocuponvigente(Me.codigocupon)
    If montocupondato = 0 Then
        MsgBox "No es posible canjear el codigo, verifique su fecha de caducidad o contactese con el area de marketing"
        Exit Sub
    Else
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, Vinculo, Observaciones)" & _
        " VALUES ('cuponx'," & montocupondato & ", " & [Forms]![Nota de Pedido]![CodPedido] & ", 0, '-', 0, '" & Me.codigocupon & "')"
        DoCmd.SetWarnings True
        MsgBox "El cupon ha sido aplicado correctamente"
        Call AplicarCuponSimple(Me.codigocupon, codigoLocal(), [Forms]![Nota de Pedido]![CodPedido])
    End If
Else
    MsgBox "No hay conexion con el servidor, contactarse con el area de sistemas"
End If
DoCmd.Close
End Sub
