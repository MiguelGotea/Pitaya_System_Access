' ==========================================================
' Modulo  : FormulasPromocionesRestricciones
' Tipo    : 1  |  Lineas: 106
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:20
' ==========================================================

Option Compare Database

Function cantidadsemillapedido(pedix As Long) As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de semillas compradas, de las grandes y de las pequenas, no galletas ni otros
miSQL = "SELECT SubPedido.CodPedido, DBBatidos.CodGrupo, [DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2 AS condi," & _
" Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, DBBatidos.CodGrupo, [DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2" & _
" HAVING (((SubPedido.CodPedido)=" & pedix & ") AND ((DBBatidos.CodGrupo)=7) AND (([DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadsemillapedido = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadsemillapedido = 0
End Function

Function cantidadsubgrupodegrupopedido(pedix As Long, grupi As Integer, subgrupi As Integer) As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de grupo y subgrupo en pedido
miSQL = "SELECT SubPedido.CodPedido, DBBatidos.CodGrupo, DBBatidos.CodSubGrupo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, DBBatidos.CodGrupo, DBBatidos.CodSubGrupo" & _
" HAVING (((SubPedido.CodPedido)=" & pedix & ") AND ((DBBatidos.CodGrupo)=" & grupi & ") AND ((DBBatidos.CodSubGrupo)=" & subgrupi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadsubgrupodegrupopedido = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadsubgrupodegrupopedido = 0
End Function


Function cantidadbatidospedido(pedix As Long) As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de semillas compradas, de las grandes y de las pequenas, no galletas ni otros
miSQL = "SELECT SubPedido.CodPedido, Grupos.Tipo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY SubPedido.CodPedido, Grupos.Tipo" & _
" HAVING (((SubPedido.CodPedido)=" & pedix & ") AND ((Grupos.Tipo)='Batido'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadbatidospedido = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadbatidospedido = 0
End Function

Function cantidadpromocionescanjeadasenbatido(pedix As Long, prom As Integer) As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad depromociones aplicadas en batidos
miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, Grupos.Tipo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion, Grupos.Tipo" & _
" HAVING (((SubPedido.CodPedido)=" & pedix & ") AND ((SubPedido.CodPromocion)=" & prom & ") AND ((Grupos.Tipo)='Batido'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadpromocionescanjeadasenbatido = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadpromocionescanjeadasenbatido = 0
End Function

Function cantidadpromocionescanjeadasensemilla(pedix As Long, prom As Integer) As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad depromociones aplicadas en batidos
miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo," & _
" [DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2 AS condi, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo, [DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2" & _
" HAVING (((SubPedido.CodPedido)=" & pedix & ") AND ((SubPedido.CodPromocion)=" & prom & ") AND ((DBBatidos.CodGrupo)=7)" & _
" AND (([DBBatidos]![CodSubGrupo]=1 Or [DBBatidos]![CodSubGrupo]=2)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadpromocionescanjeadasensemilla = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadpromocionescanjeadasensemilla = 0
End Function

