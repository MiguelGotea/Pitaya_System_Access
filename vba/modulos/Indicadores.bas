' ==========================================================
' Modulo  : Indicadores
' Tipo    : 1  |  Lineas: 621
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database

Function AcumuladoDia(afech As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT NotaDePedido.Fecha, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto" & _
" FROM NotaDePedido GROUP BY NotaDePedido.Fecha HAVING (((NotaDePedido.Fecha)=#" & afech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoDia = Round(rst("Monto"), 1)
rst.Close
     
Exit Function

Nulo:
AcumuladoDia = 0

End Function

Function AcumuladoDiaGuardado(afech As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT NotaDePedido.Fecha, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido GROUP BY NotaDePedido.Fecha HAVING (((NotaDePedido.Fecha)=#" & afech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoDiaGuardado = Round(rst("Monto"), 1)
rst.Close
     
Exit Function

Nulo:
AcumuladoDiaGuardado = 0

End Function

Function AcumuladoPeriodoGuardadoArchivoAdjunto(fdesde As Date, fhasta As Date, sucu As Integer, Tipo As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
If Tipo = "TODOS" Then
    miSQL = "SELECT [Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "# AS Expr1," & _
    " Sum(round([TotalGuardado])) AS SumaDeTotalGuardado" & _
    " FROM NotaDePedido" & sucu & "" & _
    " GROUP BY [Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#" & _
    " HAVING ((([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)<>0))"
Else
    miSQL = "SELECT [Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "# AS Expr1," & _
    " statusnotaddepedidodirecto([POS],[Delivery],[Transferencia]) AS Status," & _
    " Sum(round([TotalGuardado])) AS SumaDeTotalGuardado" & _
    " FROM NotaDePedido" & sucu & "" & _
    " GROUP BY [Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#," & _
    " statusnotaddepedidodirecto([POS],[Delivery],[Transferencia])" & _
    " HAVING ((([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)<>0)" & _
    " AND ((statusnotaddepedidodirecto([POS],[Delivery],[Transferencia]))='" & Tipo & "'))"
End If
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoPeriodoGuardadoArchivoAdjunto = Round(rst("SumaDeTotalGuardado"), 1)
rst.Close
     
Exit Function

Nulo:
AcumuladoPeriodoGuardadoArchivoAdjunto = 0

End Function



Function AcumuladoPeriodoGuardadoSinPropinaArchivoAdjunto(fdesde As Date, fhasta As Date, sucu As Integer, Tipo As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
If Tipo = "TODOS" Then
    miSQL = "SELECT ([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#) AS Expr1," & _
    " Sum(([TotalGuardado])/(1+[Propina])) AS SumaDeTotalGuardado" & _
    " FROM NotaDePedido" & sucu & "" & _
    " GROUP BY ([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)" & _
    " HAVING ((([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)<>0))"
Else

    miSQL = "SELECT ([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#) AS Expr1," & _
    " statusnotaddepedidodirecto([POS],[Delivery],[Transferencia]) AS Status," & _
    " Sum([TotalGuardado]/(1+[Propina])) AS SumaDeTotalGuardado" & _
    " FROM NotaDePedido" & sucu & "" & _
    " GROUP BY ([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)," & _
    " statusnotaddepedidodirecto([POS],[Delivery],[Transferencia])" & _
    " HAVING ((([Fecha]>=#" & fdesde & "# And [Fecha]<=#" & fhasta & "#)<>0)" & _
    " AND ((statusnotaddepedidodirecto([POS],[Delivery],[Transferencia]))='" & Tipo & "'))"
End If
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoPeriodoGuardadoSinPropinaArchivoAdjunto = Round(rst("SumaDeTotalGuardado"), 1)
rst.Close
     
Exit Function

Nulo:
AcumuladoPeriodoGuardadoSinPropinaArchivoAdjunto = 0

End Function


Function AcumuladoDiaHastaHoraGuardado(afech As Date, timex As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & timex & "# AS limitehora, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido GROUP BY NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & timex & "#" & _
" HAVING (((NotaDePedido.Fecha)=#" & afech & "#) AND (([NotaDePedido]![Hora]<#" & timex & "#)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoDiaHastaHoraGuardado = Round(rst("Monto"), 1)
rst.Close
     
Exit Function

Nulo:
AcumuladoDiaHastaHoraGuardado = 0

End Function

Function AcumuladoDiaGuardadoDesde(afech As Date, ahora As Date) As Double
On Error GoTo Nulo

Dim rst As DAO.Recordset
Dim miSQL As String
miSQL = "SELECT Sum(NotaDePedido.TotalGuardado) AS SumaDeTotalGuardado, NotaDePedido.Fecha," & _
" [NotaDePedido]![Hora]>=#" & ahora & "# AS cond" & _
" FROM NotaDePedido" & _
" GROUP BY NotaDePedido.Fecha, [NotaDePedido]![Hora]>=#" & ahora & "#" & _
" HAVING (((NotaDePedido.Fecha)=#" & afech & "#) AND (([NotaDePedido]![Hora]>=#" & ahora & "#)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

AcumuladoDiaGuardadoDesde = rst("SumaDeTotalGuardado")


rst.Close
Exit Function

Nulo:
AcumuladoDiaGuardadoDesde =  0
End Function

Function AcumuladoMesADiaGuardado(afech As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim mesa As Integer
Dim anoa As Integer

mesa = Month(afecha)
anoa = Year(afecha)
'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT Month([NotaDePedido]![Fecha])=Month(#" & afech & "#) AS condia," & _
" Year([NotaDePedido]![Fecha])=Year(#" & afech & "#) AS condim," & _
" [NotaDePedido]![Fecha]<=#" & afech & "# AS condid, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido" & _
" GROUP BY Month([NotaDePedido]![Fecha])=Month(#" & afech & "#)," & _
" Year([NotaDePedido]![Fecha])=Year(#" & afech & "#)," & _
" [NotaDePedido]![Fecha]<=#" & afech & "#" & _
" HAVING (((Month([NotaDePedido]![Fecha])=Month(#" & afech & "#))<>0)" & _
" AND ((Year([NotaDePedido]![Fecha])=Year(#" & afech & "#))<>0)" & _
" AND (([NotaDePedido]![Fecha]<=#" & afech & "#)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoMesADiaGuardado = rst("Monto")
rst.Close
     
Exit Function

Nulo:
AcumuladoMesADiaGuardado = 0

End Function

Function AcumuladoMes(ames As Integer, aano As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT Month([NotaDePedido]![Fecha]) AS Mes, Year([NotaDePedido]![Fecha]) AS Ano," & _
" Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto" & _
" FROM NotaDePedido" & _
" GROUP BY Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha])" & _
" HAVING (((Month([NotaDePedido]![Fecha]))=" & ames & ") AND ((Year([NotaDePedido]![Fecha]))=" & aano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoMes = rst("Monto")
rst.Close
     
Exit Function

Nulo:
AcumuladoMes = 0

End Function

Function AcumuladoDiaTipoVenta(fech As Date, POS As Integer, Deliv As Integer) As Double
'pos -1 = verdadero, pago con tarjeta o a cuneta
'pos= 0 = falso, pago en efectivo
'deliv segun lista
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a fecha especifica
miSQL = "SELECT NotaDePedido.Fecha, NotaDePedido.POS, NotaDePedido.Delivery, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Fecha, NotaDePedido.POS, NotaDePedido.Delivery" & _
" HAVING (((NotaDePedido.Fecha)=#" & fech & " #) AND ((NotaDePedido.POS)=" & POS & ") AND ((NotaDePedido.Delivery)=" & Deliv & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AcumuladoDiaTipoVenta = rst("Monto")
rst.Close
     
Exit Function

Nulo:
AcumuladoDiaTipoVenta = 0

End Function

Function MensualVentasGeneral(funcion As String, meta As Long, Valor As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

' funcion si es semanal, mensual o diario
' meta la meta mensual de ventas
'valor el tipo de valor que qioere 1: porcentual, 2: real

'Porcentaje de cumplimiento a la fecha y ventas hasta la fecha
miSQL = "SELECT Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto, " & funcion & "([NotaDePedido]![Fecha]) AS fecha FROM NotaDePedido GROUP BY " & funcion & "([NotaDePedido]![Fecha]) HAVING (((" & funcion & "([NotaDePedido]![Fecha]))=" & funcion & "(Date())))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Valor = 1 Then
    If funcion = "Month" Then 'Mes
        MensualVentasGeneral = rst("Monto") / (meta / diasmes(Date) * Day(Date))
        rst.Close
    
    ElseIf funcion = "numerosemana" Then 'semana
        MensualVentasGeneral = rst("Monto") / (meta / diasmes(Date) * Weekday(Date, vbMonday))
        rst.Close
    
    ElseIf funcion = "" Then ' dia
        MensualVentasGeneral = rst("Monto") / (meta / diasmes(Date))
        rst.Close
    
    End If
    
ElseIf Valor = 2 Then
    
    If funcion = "Month" Then
        MensualVentasGeneral = rst("Monto")
        rst.Close
    
    ElseIf funcion = "numerosemana" Then
        MensualVentasGeneral = rst("Monto")
        rst.Close
    
    ElseIf funcion = "" Then
        MensualVentasGeneral = rst("Monto")
        rst.Close
    
    End If

End If

Exit Function

Nulo:
MensualVentasGeneral = 0

End Function

Function MensualCantidadGeneral(funcion As String, Tipo As String, meta As Long, Valor As Integer, fechar As Date) As Double
' funcion segun mes semana o por dia
'tipo es el tipo de dato que quiero batidos bowl membresia barras etc
' valor - 1 datos porcentual valor = 2 datos reales cantidad
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'porcentaje de cumplimineto mensual en cantidad de productos segun tipo de producto y meta planteada
miSQL = "SELECT Sum(SubPedido.Cantidad) AS Total, Grupos.Tipo, " & funcion & "([NotaDePedido]![Fecha]) AS fecha" & _
" FROM Grupos INNER JOIN (DBBatidos INNER JOIN (NotaDePedido" & _
" INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo" & _
" GROUP BY Grupos.Tipo, " & funcion & "([NotaDePedido]![Fecha])" & _
" HAVING (((Grupos.Tipo)='" & Tipo & "') AND ((" & funcion & "([NotaDePedido]![Fecha]))=" & funcion & "(#" & fechar & "#)))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Valor = 1 Then
If funcion = "Month" Then
MensualCantidadGeneral = rst("Total") / (meta / diasmes(Date) * Day(Date))
rst.Close

ElseIf funcion = "numerosemana" Then
MensualCantidadGeneral = rst("Total") / (meta / diasmes(Date) * Weekday(Date, vbMonday))
rst.Close

ElseIf funcion = "" Then
MensualCantidadGeneral = rst("Total") / CInt((meta / diasmes(Date)))
rst.Close

End If

ElseIf Valor = 2 Then

If funcion = "Month" Then
MensualCantidadGeneral = rst("Total")
rst.Close

ElseIf funcion = "numerosemana" Then
MensualCantidadGeneral = rst("Total")
rst.Close

ElseIf funcion = "" Then
MensualCantidadGeneral = rst("Total")
rst.Close

End If

End If


' Meta Mensual en cantidad


Exit Function

Nulo:
MensualCantidadGeneral = 0

End Function

Function PorcentajeUsoVidrio(Fecha As Date) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim vidrio As Double

'Cantidad de batidos en plastico entre total vendido
miSQL = "SELECT NotaDePedido.Fecha, Grupos.Tipo, Sum(SubPedido.Cantidad) AS Total" & _
" FROM Grupos INNER JOIN (DBBatidos INNER JOIN (NotaDePedido" & _
" INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo" & _
" GROUP BY NotaDePedido.Fecha, Grupos.Tipo, NotaDePedido.Modalidad HAVING (((NotaDePedido.Fecha)=Date())" & _
" AND ((Grupos.Tipo)='Batido') AND ((NotaDePedido.Modalidad)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PorcentajeUsoVidrio = rst("Total") / MensualCantidadGeneral("", "Batido", 0, 2, Fecha)
rst.Close

Exit Function

Nulo:
PorcentajeUsoVidrio = 0

End Function

Function PorcentajeMedidasBatidos(tipof As String, fechax As Date) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim vidrio As Double

'Cantidad de batidos en plastico entre total vendido
miSQL = "SELECT DBBatidos.Medida, NotaDePedido.Fecha, NotaDePedido.Anulado, Sum(SubPedido.Cantidad) AS Total" & _
" FROM NotaDePedido INNER JOIN (DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido)" & _
" ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY DBBatidos.Medida, NotaDePedido.Fecha, NotaDePedido.Anulado" & _
" HAVING (((DBBatidos.Medida)='" & tipof & "') AND ((NotaDePedido.Fecha)=#" & fechax & "#) AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PorcentajeMedidasBatidos = rst("Total") / CantidadVentasGrupoTamanoDia(fechax, tipof)
rst.Close

Exit Function

Nulo:
PorcentajeMedidasBatidos = 0

End Function

Function CantidadVentasGrupoTamanoDia(adia As Date, taman As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

Dim rst2 As DAO.Recordset
Dim miSQL2 As String

Dim rst3 As DAO.Recordset
Dim miSQL3 As String

Dim filtro1 As String
Dim filtro2 As String
Dim filtro3 As String
Dim q1 As Integer
Dim q2 As Integer
Dim q3 As Integer

'Sacamos los grupos que contienen ese tamano
miSQL = "SELECT DBBatidos.Medida, Grupos.Tipo FROM DBBatidos INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY DBBatidos.Medida, Grupos.Tipo HAVING (((DBBatidos.Medida)='" & taman & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)


rst.MoveLast
q1 = rst.RecordCount
rst.MoveFirst

filtro1 = ""

For I = 1 To q1
    filtro1 = filtro1 & "(Grupos.Tipo)='" & rst("Tipo") & "' Or "
    rst.MoveNext
Next I
filtro1 = Left(filtro1, Len(filtro1) - 4)
rst.Close

'Sacamaos los tamanos que contienn todos los grupos de la anteior consulta
miSQL2 = "SELECT DBBatidos.Medida, Grupos.Tipo FROM DBBatidos INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY DBBatidos.Medida, Grupos.Tipo HAVING ((" & filtro1 & "))"
Set rst2 = CurrentDb.OpenRecordset(miSQL2, dbOpenDynaset)

rst2.MoveLast
q2 = rst2.RecordCount
rst2.MoveFirst

filtro2 = ""

For I = 1 To q2
    filtro2 = filtro2 & "(DBBatidos.Medida)='" & rst2("Medida") & "' Or "
    rst2.MoveNext
Next I
filtro2 = Left(filtro2, Len(filtro2) - 4)
rst2.Close

'Ventas de todos los tamanos del filtro anteior
miSQL3 = "SELECT DBBatidos.Medida, NotaDePedido.Fecha, NotaDePedido.Anulado, Sum(SubPedido.Cantidad) AS Total" & _
" FROM NotaDePedido INNER JOIN (DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido)" & _
" ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY DBBatidos.Medida, NotaDePedido.Fecha, NotaDePedido.Anulado" & _
" HAVING ((" & filtro2 & ") AND ((NotaDePedido.Fecha)=#" & adia & "#) AND ((NotaDePedido.Anulado)=0))"

Set rst3 = CurrentDb.OpenRecordset(miSQL3, dbOpenDynaset)

rst3.MoveLast
q3 = rst3.RecordCount
rst3.MoveFirst
CantidadVentasGrupoTamanoDia = 0

For I = 1 To q3
    CantidadVentasGrupoTamanoDia = CantidadVentasGrupoTamanoDia + rst3("Total")
    rst3.MoveNext
Next I
rst3.Close

Exit Function

Nulo:
CantidadVentasGrupoTamanoDia = 0
End Function

Function CantidadVentasTipoDia(adia As Date, tipox As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Anulado, Grupos.Tipo, NotaDePedido.Fecha, Sum([SubPedido]![Cantidad]) AS Cantidad" & _
" FROM ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY NotaDePedido.Anulado, Grupos.Tipo, NotaDePedido.Fecha" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((Grupos.Tipo)='" & tipox & "') AND ((NotaDePedido.Fecha)=#" & adia & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadVentasTipoDia = rst("Cantidad")
rst.Close

Exit Function

Nulo:
CantidadVentasTipoDia = 0

End Function

Function CantidadVentasProductoxCodigoDia(adia As Date, produ As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, SubPedido.CodBatido, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, SubPedido.CodBatido" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & adia & "#)" & _
" AND ((SubPedido.CodBatido)='" & produ & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadVentasProductoxCodigoDia = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
CantidadVentasProductoxCodigoDia = 0

End Function

Function CantidadVentasProductoxNombreDia(adia As Date, produ As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, DBBatidos.Nombre, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, DBBatidos.Nombre" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & adia & "#)" & _
" AND ((DBBatidos.Nombre)='" & produ & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadVentasProductoxNombreDia = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
CantidadVentasProductoxNombreDia = 0

End Function

Function ConsumoPorPorciones(ING As String, Porcion As Double, tipoprod As String) As Long
' historial de 2 semanas, semana anterior y una antes

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de consumo en semana especifica de ingrediente acorde a las ventas de dos semanas anteriores
miSQL = "SELECT mcm([SubReceta]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],0) AS Cantidad, Sum(0.5*[SubPedido]![Cantidad]*mcm([SubReceta]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],1)) AS Total, SubReceta.CodIngrediente, Grupos.Tipo " & _
"FROM Grupos INNER JOIN (((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo GROUP BY mcm([SubReceta]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],0), SubReceta.CodIngrediente, Grupos.Tipo, numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-3, numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()) " & _
"HAVING (((mcm([SubReceta]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],0))=" & Porcion & ") AND ((SubReceta.CodIngrediente)='" & ING & "') AND ((Grupos.Tipo)='" & tipoprod & "') AND ((numerosemana([NotaDePedido]![Fecha])>numerosemana(Date())-3)<>0) AND ((numerosemana([NotaDePedido]![Fecha])<numerosemana(Date()))<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoPorPorciones = rst("Total")
rst.Close

Exit Function

Nulo:
ConsumoPorPorciones = 0

End Function

Function PorcentajeAgua(bat As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Med As String

'Hallar la medida del batido
miSQL = "SELECT DBBatidos.CodBatido, DBBatidos.Medida FROM DBBatidos WHERE (((DBBatidos.CodBatido)='" & bat & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Med = rst("Medida")
rst.Close

If Med = "Kid" Then
    PorcentajeAgua = (CantidadIngrediente(bat, "B004") + CantidadIngrediente(bat, "B001")) / 340

ElseIf Med = "Mediano" Then
    PorcentajeAgua = (CantidadIngrediente(bat, "B004") + CantidadIngrediente(bat, "B001")) / 454

ElseIf Med = "Gigantona" Then
    PorcentajeAgua = (CantidadIngrediente(bat, "B004") + CantidadIngrediente(bat, "B001")) / 624

Else
    PorcentajeAgua = 0

End If
     
Exit Function

Nulo:
PorcentajeAgua = 0

End Function

Function MetaMensual(Tipo As String, Mes As Integer) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de la meta del mes especifico del tipo de producto especifico
miSQL = "SELECT MetasMensuales.Mes, MetasMensuales.[" & Tipo & "] FROM MetasMensuales WHERE (((MetasMensuales.Mes)=" & Mes & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MetaMensual = rst("" & Tipo & "")
rst.Close

Exit Function

Nulo:
MetaMensual = 0

End Function
