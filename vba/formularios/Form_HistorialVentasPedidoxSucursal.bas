' ==========================================================
' Modulo  : Form_HistorialVentasPedidoxSucursal
' Tipo    : 100  |  Lineas: 43
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database


Public Sub Comando2291_Click()

Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "NotaDePedido", "NotaDePedidoRevision", 1)
'Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "NotaDePedido", "NotaDePedido", 1)
If Me.atipo = "TODOS" Then
    If Me.stipo = "TODOS" Then
        Me.Filter = "[Fecha]>= #" & Me.fechadesde & "# And [Fecha]<= #" & Me.fechahasta & "#"
        Me.FilterOn = True
    Else
        Me.Filter = "[Fecha]>= #" & Me.fechadesde & "# And [Fecha]<= #" & Me.fechahasta & "# And [Estado]='" & Me.stipo & "'"
        Me.FilterOn = True
    End If
Else
    If Me.stipo = "TODOS" Then
        Me.Filter = "[Fecha]>= #" & Me.fechadesde & "# And [Fecha]<= #" & Me.fechahasta & "# And [Pago]='" & Me.atipo & "'"
        Me.FilterOn = True
    Else
        Me.Filter = "[Fecha]>= #" & Me.fechadesde & "# And [Fecha]<= #" & Me.fechahasta & "# And [Pago]='" & Me.atipo & "' And [Estado]='" & Me.stipo & "'"
        Me.FilterOn = True
    End If
End If
Me.Requery
End Sub

Private Sub Form_Close()
Me.RecordSource = ""
DoCmd.RunSQL "DROP TABLE NotaDePedidoRevision"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub







