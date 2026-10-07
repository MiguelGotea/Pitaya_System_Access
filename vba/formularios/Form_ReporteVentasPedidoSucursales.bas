' ==========================================================
' Modulo  : Form_ReporteVentasPedidoSucursales
' Tipo    : 100
' Lineas  : 51
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database


Public Sub Comando2291_Click()
If Me.asucursal <> 0 Then

    Dim rst As DAO.Recordset
    Dim miSQL As String
    miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
    " FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    Do While Not rst.EOF
        Call importartablaespecifica("Pitaya" & rst("CodLocal") & "_DB", "NotaDePedido", "NotaDePedido" & rst("CodLocal"), 3)
        rst.MoveNext
    Loop
    rst.Close

Else
    Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "NotaDePedido", "NotaDePedido" & Me.asucursal, 3)
End If
Me.Requery

End Sub

Private Sub Form_Close()
Dim rst As DAO.Recordset
Dim miSQL As String
miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    If DCount("*", "MSysObjects", "Type=1 AND Name='NotaDePedido' & " & rst("CodLocal")) > 0 Then
        DoCmd.RunSQL "DROP TABLE NotaDePedido" & rst("CodLocal")
    End If
    
    rst.MoveNext
Loop
rst.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub







