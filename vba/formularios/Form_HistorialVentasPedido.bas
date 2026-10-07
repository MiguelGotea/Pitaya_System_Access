' ==========================================================
' Modulo  : Form_HistorialVentasPedido
' Tipo    : 100
' Lineas  : 42
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database

Private Sub Comando2291_Click()

If Me.atipo = "TODOS" Then
    If Me.stipo = "TODOS" Then
        Me.Filter = "[Fecha]=#" & Me.fechasistema & "#"
        Me.FilterOn = True
    Else
        Me.Filter = "[Fecha]=#" & Me.fechasistema & "# And [Estado]='" & Me.stipo & "'"
        Me.FilterOn = True
    End If
Else
    If Me.stipo = "TODOS" Then
        Me.Filter = "[Fecha]=#" & Me.fechasistema & "# And [Pago]='" & Me.atipo & "'"
        Me.FilterOn = True
    Else
        Me.Filter = "[Fecha]=#" & Me.fechasistema & "# And [Pago]='" & Me.atipo & "' And [Estado]='" & Me.stipo & "'"
        Me.FilterOn = True
    End If
End If
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Me.Filter = "[Fecha]=#" & Date & "#"
Me.FilterOn = True

End Sub



Private Sub Comando50_Click()
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Me.CodPedido
End Sub




