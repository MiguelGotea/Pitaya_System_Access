' ==========================================================
' Modulo  : Form_ValidacionCedulaClub
' Tipo    : 100  |  Lineas: 36
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:20
' ==========================================================

Option Compare Database

Private Sub Comando65_Click()
Dim cedulahost As String
cedulahost = verificarcedulaclub(Me.codigoclub)
If cedulahost = "sin_registro" Then
    If EsSistemaDeTienda() Then
        Call guardarnumerocedulahost(Me.codigoclub, Me.cedulaingresada)
        cedulahost = verificarcedulaclub(Me.codigoclub)
    End If
Else
    If cedulahost = "sin_registro" Then
        MsgBox "No existe cedula registrada del cliente, solicitar a Marketing completar el registro"
        Call resetearpromocionespedidocompleto(Me.codigopedido)
        [Forms]![Nota de Pedido].Form.ppedido.Requery
        [Forms]![Nota de Pedido].Form.Texto193.Requery
    Else
        If Me.cedulaingresada <> cedulahost Then
            MsgBox "La cedula ingresada no coincide con la cedula registrada, solicitar a Marketing la revision"
            Call resetearpromocionespedidocompleto(Me.codigopedido)
            [Forms]![Nota de Pedido].Form.ppedido.Requery
            [Forms]![Nota de Pedido].Form.Texto193.Requery
        Else
            MsgBox "Cedula Validada"
            [Forms]![Nota de Pedido].VALIDACIONCEDULA = True
        End If
    End If

End If
DoCmd.Close acForm, "ValidacionCedulaClub"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
