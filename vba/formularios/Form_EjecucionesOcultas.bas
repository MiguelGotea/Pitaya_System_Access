' ==========================================================
' Modulo  : Form_EjecucionesOcultas
' Tipo    : 100  |  Lineas: 36
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================


Option Compare Database

Private Sub Comando0_Click()
On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String
Dim codpedi As Long
Dim codsubpedi As Long
Dim cont As Long
Dim ex As Long

miSQL = "SELECT SubPedido.CodSubPedido, NotaDePedido.CodPedido, Year([NotaDePedido]![Fecha]) AS Expr1, NotaDePedido.CodCliente" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" WHERE (((Year([NotaDePedido]![Fecha]))<2024) AND ((NotaDePedido.CodCliente)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cont = rst.RecordCount
rst.MoveFirst

For ex = 1 To cont
    codpedi = rst("CodPedido")
    codsubpedi = rst("CodSubPedido")
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM SubPedido WHERE CodSubPedido =" & codsubpedi
    DoCmd.RunSQL "DELETE * FROM NotaDePedido WHERE CodPedido =" & codpedi
    DoCmd.SetWarnings True
    
    rst.MoveNext

Next ex

MsgBox "fin"

End Sub
